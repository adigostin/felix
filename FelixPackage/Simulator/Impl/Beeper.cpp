
#include "pch.h"
#include "Bus.h"
#include <xaudio2.h>

class Beeper : public IDevice, IXAudio2VoiceCallback
{
	Bus* _io_bus;

	bool _level = false;
	UINT64 _time = 0;

	static constexpr uint32_t osc_freq = 3'500'000;
	static constexpr uint32_t sample_freq = 35000;
	static constexpr uint8_t bits_per_sample = 8;
	static constexpr uint32_t max_delay_ms = 20;
	static constexpr uint32_t max_delay_t_states = osc_freq * max_delay_ms / 1000;
	static constexpr uint32_t buffer_length_samples = sample_freq * max_delay_ms / 1000;

	vector_nothrow<uint8_t> _samples;

	// This is meaningful only when _samples is empty.
	bool _previous_packet_last_sample_level;
	
	// this is zero at the beginning, and also after each timeout between packets
	UINT64 _previous_packet_last_sample_time = 0;

	wil::com_ptr_nothrow<IXAudio2> _xaudio2;
	IXAudio2MasteringVoice* _mastering_voice = nullptr;
	IXAudio2SourceVoice* _source_voice = nullptr;

public:
	HRESULT InitInstance (Bus* io_bus)
	{
		_io_bus = io_bus;

		// Let's allocate upfront the space we need.
		bool reserved = _samples.try_reserve(buffer_length_samples * bits_per_sample / 8); RETURN_HR_IF(E_OUTOFMEMORY, !reserved);

		bool pushed = _io_bus->write_responders.try_push_back({ this, &process_io_write_request }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);

		auto hr = XAudio2Create (&_xaudio2, 0, XAUDIO2_USE_DEFAULT_PROCESSOR); RETURN_IF_FAILED(hr);
		
		//XAUDIO2_DEBUG_CONFIGURATION xadc = { };
		//xadc.TraceMask = XAUDIO2_LOG_WARNINGS | XAUDIO2_LOG_DETAIL;
		//xadc.BreakMask = XAUDIO2_LOG_WARNINGS;
		//_xaudio2->SetDebugConfiguration(&xadc);

		uint32_t sound_channel_count = 1;

		hr = _xaudio2->CreateMasteringVoice(&_mastering_voice, sound_channel_count, sample_freq); RETURN_IF_FAILED(hr);
		auto destroyMasteringVoice = wil::scope_exit([this] { _mastering_voice->DestroyVoice(); _mastering_voice = nullptr; });

		WAVEFORMATEX wfx;
		wfx.wFormatTag = WAVE_FORMAT_PCM;
		wfx.nChannels = 1; // mono (not stereo, not 5+1 or something)
		wfx.nSamplesPerSec = sample_freq;
		wfx.nAvgBytesPerSec = sample_freq * bits_per_sample / 8;
		wfx.nBlockAlign = wfx.nChannels * bits_per_sample / 8;
		wfx.wBitsPerSample = bits_per_sample;
		wfx.cbSize = 0;
		hr = _xaudio2->CreateSourceVoice (&_source_voice, &wfx, 0, 2.0f, this); RETURN_IF_FAILED(hr);
		auto destroySourceVoice = wil::scope_exit([this] { _source_voice->DestroyVoice(); _source_voice = nullptr; });

		hr = _source_voice->Start(0); RETURN_IF_FAILED(hr);

		destroySourceVoice.release();
		destroyMasteringVoice.release();

		return S_OK;
	}

	~Beeper()
	{
		_source_voice->DestroyVoice(); _source_voice = nullptr;
		_mastering_voice->DestroyVoice(); _mastering_voice = nullptr;
	}

	#pragma region IDevice
	virtual void STDMETHODCALLTYPE Reset() override
	{
		_time = 0;
		_level = 0;
		_samples.clear();
		_previous_packet_last_sample_level = false;
		_previous_packet_last_sample_time = 0;
		uint8_t data = 0;
		AudioBufferEvent(&data, 1);
	}

	virtual UINT64 STDMETHODCALLTYPE Time() override { return _time; }

	virtual BOOL STDMETHODCALLTYPE NeedSyncWithRealTime (UINT64* sync_time) override { return FALSE; }

	virtual bool SimulateTo (UINT64 requested_time) override
	{
		WI_ASSERT (_time < requested_time);

		auto initial_time = _time;

		constexpr uint32_t increment = osc_freq / sample_freq;
		static_assert(osc_freq % sample_freq == 0);

		while (_time < requested_time)
		{
			if (_io_bus->writer_behind_of(_time))
				return S_FALSE;

			// TODO: after a long pause, send a first buffer with double the normal size and double the normal delay.
			// If we send this first buffer with the normal size and the normal delay,
			// the slighest delay between subsequent buffers will cause the audio to interrupt.

			if (_samples.empty())
			{
				// We start a new packet only if the level changed from the last sample in the previous packet.
				if (_level != _previous_packet_last_sample_level)
				{
					// The level did change.

					// Was that previous packet in the recent past, or a long time ago?
					if (!_previous_packet_last_sample_time)
					{
						// It was a long time ago. (We erased its timestamp to avoid overflowing time deltas.)
						_samples.try_push_back(_level ? 127 : 0);
					}
					else
					{
						// Ok, it was in the recent past.
						// That previous packet should not have been too long ago.
						WI_ASSERT (_time - _previous_packet_last_sample_time < max_delay_t_states);

						// Here we generate samples between the last sample of the previous packet and the sample we have now.
						for (auto t = _previous_packet_last_sample_time + increment; (int32_t)(_time - t) > 0; t += increment)
						{
							WI_ASSERT(_samples.size() < _samples.capacity());
							_samples.try_push_back(_previous_packet_last_sample_level ? 127 : 0);
						}

						// and now we add the current sample
						WI_ASSERT(_samples.size() < _samples.capacity());
						_samples.try_push_back(_level ? 127 : 0);

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
				WI_ASSERT(_samples.size() < _samples.capacity());
				_samples.try_push_back(_level ? 127 : 0);

				if (_samples.size() == _samples.capacity())
				{
					AudioBufferEvent (_samples.data(), _samples.size());
					_previous_packet_last_sample_level = _level;
					_previous_packet_last_sample_time = _time;
					_samples.clear();
				}
			}

			_time += increment;
		}

		return true;
	}
	#pragma endregion

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

	void AudioBufferEvent (const uint8_t* data, uint32_t data_size_bytes)
	{
		XAUDIO2_VOICE_STATE state;
		_source_voice->GetState (&state);
		if (state.BuffersQueued == XAUDIO2_MAX_QUEUED_BUFFERS)
		{
			// At the time of this writing, this happens when the processor is starved, and can be easily reproduced
			// by running in the simulator SAVE "D" CODE 0,10000 while running something like HeavyLoad,
			// on all processor cores, with Above Normal priority, for about 30 seconds; when stopping HeavyLoad,
			// the simulator tries to catch up (bad idea - needs fixing), so it generates many audio buffers.
			return;
		}

		auto copy = (uint8_t*)malloc (data_size_bytes);
		if (!copy)
			return;
		memcpy (copy, data, data_size_bytes);

		XAUDIO2_BUFFER buffer = { };
		buffer.pAudioData = copy;
		buffer.AudioBytes = data_size_bytes;
		buffer.pContext = copy;
		auto hr = _source_voice->SubmitSourceBuffer (&buffer);
		if (FAILED(hr))
			// Not sure how to handle this. We're not on the GUI thread here so our telemetry dialog code will probably crash.
			free(copy);
	}

	#pragma region IXAudio2VoiceCallback
	virtual void STDMETHODCALLTYPE OnVoiceProcessingPassStart (UINT32 BytesRequired) override
	{
	}

	virtual void STDMETHODCALLTYPE OnVoiceProcessingPassEnd() override
	{
	}

	virtual void STDMETHODCALLTYPE OnStreamEnd (THIS) override
	{
	}

	virtual void STDMETHODCALLTYPE OnBufferStart (void* pBufferContext) override
	{
	}

	virtual void STDMETHODCALLTYPE OnBufferEnd (void* pBufferContext) override
	{
		free (pBufferContext);
	}

	virtual void STDMETHODCALLTYPE OnLoopEnd (void* pBufferContext) override
	{
	}

	virtual void STDMETHODCALLTYPE OnVoiceError (void* pBufferContext, HRESULT Error) override
	{
	}
	#pragma endregion
};

HRESULT STDMETHODCALLTYPE MakeBeeper (Bus* io_bus, wistd::unique_ptr<IDevice>* ppDevice)
{
	auto d = wil::make_unique_nothrow<Beeper>(); RETURN_IF_NULL_ALLOC(d);
	auto hr = d->InitInstance(io_bus); RETURN_IF_FAILED(hr);
	*ppDevice = std::move(d);
	return S_OK;
}
