
#include "pch.h"
#include "SimulatorInternal.h"

class TapPlayer : public ITapPlayerDevice
{
	Bus* _io_bus;
	ITapPlayerEventHandler* _eh;
	bool _iolevel = false;
	uint32_t _index;
	vector_nothrow<tap_block_t> _queued_blocks;

	uint64_t _prev_pulse_time;
	vector_nothrow<uint16_t> _pulse_lengths;
	uint32_t _next_pulse_index;
	uint32_t _silence_after;

	static_assert(osc_freq % sample_freq == 0);

	uint8_t audio[buffer_length_samples * bits_per_sample / 8];
	uint32_t audio_size;
	uint64_t audio_next_sample_time;
	IXAudio2SourceVoice* _source_voice = nullptr;
	XAudio2VoiceCallback callback;

public:
	HRESULT InitInstance (Bus* io_bus, IXAudio2* xaudio2, ITapPlayerEventHandler* eh)
	{
		HRESULT hr;
		_io_bus = io_bus;
		_eh = eh;
		bool pushed = io_bus->read_responders.try_push_back({ this, &ProcessIoReadRequest }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		
		static const WAVEFORMATEX wfx = {
			.wFormatTag = WAVE_FORMAT_PCM,
			.nChannels = 1, // mono (not stereo, not 5+1 or something)
			.nSamplesPerSec = sample_freq,
			.nAvgBytesPerSec = sample_freq * bits_per_sample / 8,
			.nBlockAlign = /*wfx.nChannels*/1 * bits_per_sample / 8,
			.wBitsPerSample = bits_per_sample,
			.cbSize = 0,
		};

		hr = xaudio2->CreateSourceVoice (&_source_voice, &wfx, 0, 2.0f, &callback); RETURN_IF_FAILED(hr);
		hr = _source_voice->Start(0); RETURN_IF_FAILED(hr);

		return S_OK;
	}

	#pragma region IDevice
	virtual void Reset() override
	{
		_time = 0;
		_iolevel = false;
		StopPlaying();
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return FALSE; }

	void PushSample (uint8_t audio_level)
	{
		WI_ASSERT(audio_size < _countof(audio));
		audio[audio_size++] = audio_level;
		if (audio_size == _countof(audio))
		{
			SendSamplesToXAudio (_source_voice, audio, audio_size);
			audio_size = 0;
		}
	}

	virtual void SimulateDeviceTo (UINT64 requested_time) override
	{
		if (_pulse_lengths.empty())
		{
			// Nothing playing.
			_time = requested_time;
			return;
		}
		
		while(true)
		{
			if (_next_pulse_index == 0)
			{
				WI_ASSERT(_pulse_lengths[0] == 0);
				_iolevel = true;
				_prev_pulse_time = _time;
				_next_pulse_index = 1;
				WI_ASSERT(audio_size == 0);
				audio[0] = audio_level_high;
				audio_size = 1;
				audio_next_sample_time = _time + audio_increment;

			}
			else if (_next_pulse_index < _pulse_lengths.size())
			{
				auto next_pulse_time = _prev_pulse_time + _pulse_lengths[_next_pulse_index];
				if (next_pulse_time > requested_time)
				{
					while (audio_next_sample_time + audio_increment <= requested_time)
					{
						PushSample(_iolevel ? audio_level_high : audio_level_low);
						audio_next_sample_time += audio_increment;
					}

					_time = requested_time;
					break;
				}

				// next_pulse_time <= requested_time

				while (audio_next_sample_time + audio_increment < next_pulse_time)
				{
					PushSample(_iolevel ? audio_level_high : audio_level_low);
					audio_next_sample_time += audio_increment;
				}

				_time = next_pulse_time;
				_prev_pulse_time = next_pulse_time;
				_iolevel = !_iolevel;
				_next_pulse_index++;

				if (audio_next_sample_time + audio_increment == next_pulse_time)
				{
					PushSample(_iolevel ? audio_level_high : audio_level_low);
					audio_next_sample_time += audio_increment;
				}

				if (_next_pulse_index == _pulse_lengths.size())
				{
					// silence begins
					audio[0] = audio_level_silence;
					SendSamplesToXAudio (_source_voice, audio, 1);
					audio_size = 0;
				}
			}
			else
			{
				// silence period
				auto next_block_time = _prev_pulse_time + _silence_after;
				if (next_block_time > requested_time)
				{
					_time = requested_time;
					break;
				}

				// silence period ends
				_pulse_lengths.clear();
				_next_pulse_index = 0;
				if (_queued_blocks.empty())
				{
					_time = requested_time;
					_eh->OnTapPlayComplete();
					break;
				}

				_time = next_block_time;
				GeneratePulses();
			}
		}
	}
	#pragma endregion

	#pragma region ITapPlayerDevice
	virtual HRESULT STDMETHODCALLTYPE AddBlocks (vector_nothrow<tap_block_t> blocks) override
	{
		if (!_queued_blocks.empty())
			return E_UNEXPECTED;
		if (!_pulse_lengths.empty())
			return E_UNEXPECTED;
		_queued_blocks = std::move(blocks);
		_eh->OnTapPlayStarting();
		return GeneratePulses();
	}

	virtual HRESULT STDMETHODCALLTYPE StopPlaying() override
	{
		if (_pulse_lengths.size())
		{
			// currently playing
			_pulse_lengths.clear();
			_queued_blocks.clear();

			audio[0] = audio_level_silence;
			SendSamplesToXAudio (_source_voice, audio, 1);
			audio_size = 0;

			_eh->OnTapPlayComplete();
		}

		return S_OK;
	}
	#pragma endregion

	static uint8_t ProcessIoReadRequest (IDevice* device, uint16_t address)
	{
		if ((address & 0xFF) == 0xFE)
		{
			// keys and Tape In
			auto ai = static_cast<TapPlayer*>(device);
			return ai->_iolevel ? 0xBF : 0xFF;
		}

		return 0xFF;
	}

	HRESULT GeneratePulses()
	{
		auto block = _queued_blocks.remove(_queued_blocks.begin());
		_pulse_lengths.clear();
		size_t len = block.get()[0] | (block.get()[1] << 8);

		// https://sinclair.wiki.zxnet.co.uk/wiki/Spectrum_tape_interface

		_prev_pulse_time = _time;
		_pulse_lengths.try_push_back(0);

		// Pilot
		bool isHeader = block.get()[2] == 0;
		size_t pulseCount = isHeader ? 8063 : 3223;
		for (size_t i = 0; i < pulseCount; i++)
			_pulse_lengths.try_push_back(2168);

		// Sync
		_pulse_lengths.try_push_back(667);
		_pulse_lengths.try_push_back(735);

		// Data
		uint8_t* p = block.get() + 2;
		uint8_t* to = p + len;
		while(p < to)
		{
			uint8_t b = *p;
			for (uint8_t m = 0x80; m; m >>= 1)
			{
				uint16_t l = (b & m) ? 1710 : 855;
				_pulse_lengths.try_push_back(l);
				_pulse_lengths.try_push_back(l);
			}
			p++;
		}

		_silence_after = 3'500'000;

		_next_pulse_index = 0;

		return S_OK;
	}
};

HRESULT STDMETHODCALLTYPE MakeTapPlayer (Bus* io_bus, IXAudio2* xaudio2, ITapPlayerEventHandler* eh, wistd::unique_ptr<ITapPlayerDevice>& ppDevice)
{
	auto d = wil::make_unique_nothrow<TapPlayer>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(io_bus, xaudio2, eh); RETURN_IF_FAILED(hr);
	ppDevice = std::move(d);
	return S_OK;
}

