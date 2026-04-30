
#include "pch.h"
#include "SimulatorInternal.h"

class Beeper : public IDevice
{
	static constexpr uint8_t bits_per_sample = 8;

	bool _level = false;

	uint8_t data[buffer_length_samples * bits_per_sample / 8];
	size_t size = 0;

	// This is meaningful only when size is 0.
	bool _previous_packet_last_sample_level;

	// this is zero at the beginning, and also after each timeout between packets
	UINT64 _previous_packet_last_sample_time = 0;

	IXAudio2SourceVoice* _source_voice = nullptr;
	XAudio2VoiceCallback callback;

public:
	HRESULT InitInstance (Bus* io_bus, IXAudio2* xaudio2)
	{
		HRESULT hr;
		bool pushed = io_bus->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

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

	~Beeper()
	{
		if (_source_voice)
		{
			_source_voice->DestroyVoice();
			_source_voice = nullptr;
		}
	}

	#pragma region IDevice
	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
		_level = 0;
		size = 0;
		_previous_packet_last_sample_level = false;
		_previous_packet_last_sample_time = 0;
		data[0] = audio_level_silence;
		SendSamplesToXAudio (_source_voice, data, 1);
	}

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return FALSE; }

	virtual void SimulateTo (UINT64 requested_time) override
	{
		WI_ASSERT (_time < requested_time);

		static_assert(osc_freq % sample_freq == 0);

		while (_time < requested_time)
		{
			// TODO: after a long pause, send a first buffer with double the normal size and double the normal delay.
			// If we send this first buffer with the normal size and the normal delay,
			// the slighest delay between subsequent buffers will cause the audio to interrupt.

			if (size == 0)
			{
				// We start a new packet only if the level changed from the last sample in the previous packet.
				if (_level != _previous_packet_last_sample_level)
				{
					// The level did change.

					// Was that previous packet in the recent past, or a long time ago?
					if (!_previous_packet_last_sample_time)
					{
						// It was a long time ago.
						PushSample(_level ? 127 : 0);
					}
					else
					{
						// Ok, it was in the recent past.
						// That previous packet should not have been too long ago.
						WI_ASSERT (_time - _previous_packet_last_sample_time <= max_delay_t_states);

						// Here we generate samples between the last sample of the previous packet and the sample we have now.
						for (auto t = _previous_packet_last_sample_time + audio_increment; (int32_t)(_time - t) > 0; t += audio_increment)
							PushSample(_previous_packet_last_sample_level ? 127 : 0);

						// and now we add the current sample
						PushSample(_level ? 127 : 0);

						// From now on we should have no need to know about the last sample of the previous packet.
						_previous_packet_last_sample_time = 0;
					}
				}
				else
				{
					// The level did not change compared to the last sample of the previous packet.
					if (_previous_packet_last_sample_time
						&& (_time - _previous_packet_last_sample_time >= max_delay_t_states))
					{
						// A long time has passed since the previous packet. Let's reset the information about
						// that previous packet, or else the time offset from it might grow to be larger than 32 bits.
						_previous_packet_last_sample_time = 0;
					}
				}
			}
			else
			{
				// add current level as sample to the existing packet
				PushSample(_level ? 127 : 0);
			}

			_time += audio_increment;
		}
	}
	#pragma endregion

	void PushSample (uint8_t value)
	{
		WI_ASSERT(size < _countof(data));
		data[size++] = value;

		if (size == _countof(data))
		{
			SendSamplesToXAudio (_source_voice, data, size);
			_previous_packet_last_sample_level = _level;
			_previous_packet_last_sample_time = _time;
			size = 0;
		}
	}

	static void process_io_write_request (IDevice* d, uint16_t address, uint8_t value)
	{
		if ((address & 0xFF) == 0xFE)
		{
			auto* b = static_cast<Beeper*>(d);
			bool new_level = !!(value & 0x10) ^ !!(value & 8);
			if (b->_level != new_level)
			{
				b->_level = new_level;
				//OutputDebugString(new_level ? L"1" : L"0");
			}
		}
	}
};

HRESULT STDMETHODCALLTYPE MakeBeeper (Bus* io_bus, IXAudio2* xaudio2, wistd::unique_ptr<IDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<Beeper>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(io_bus, xaudio2); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
