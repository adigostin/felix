#include "pch.h"
#include "SimulatorInternal.h"
#include <algorithm>

class AYChip : public IDevice
{
	Bus* _ioBus = nullptr;
	IXAudio2SourceVoice* _sourceVoice = nullptr;
	XAudio2VoiceCallback _callback;
	SpectrumVariant _variant;
	uint8_t _registers[16] = {};
	uint8_t _selectedRegister = 0;
	uint8_t _samples[buffer_length_samples];
	size_t _sampleCount = 0;
	uint32_t _toneCounters[3] = {};
	bool _toneLevels[3] = {};
	uint32_t _noiseCounter = 0;
	uint32_t _noiseShiftRegister = 0x1FFFF;
	bool _noiseLevel = true;
	uint32_t _envelopeCounter = 0;
	uint8_t _envelopeCount = 15;
	uint8_t _envelopeAttackMask = 0;
	bool _envelopeHolding = false;

	static constexpr uint8_t volumeLevels[16] = {
		0, 1, 1, 1, 2, 2, 3, 4, 5, 7, 9, 12, 16, 21, 28, 36
	};

public:
	HRESULT InitInstance(Bus* ioBus, IXAudio2* xaudio2, SpectrumVariant variant)
	{
		_ioBus = ioBus;
		_variant = variant;

		static const WAVEFORMATEX wfx = {
			.wFormatTag = WAVE_FORMAT_PCM,
			.nChannels = 1,
			.nSamplesPerSec = sample_freq,
			.nAvgBytesPerSec = sample_freq * bits_per_sample / 8,
			.nBlockAlign = 1 * bits_per_sample / 8,
			.wBitsPerSample = bits_per_sample,
			.cbSize = 0,
		};

		HRESULT hr = xaudio2->CreateSourceVoice(&_sourceVoice, &wfx, 0, 2.0f, &_callback); RETURN_IF_FAILED(hr);
		hr = _sourceVoice->Start(0); RETURN_IF_FAILED(hr);

		bool pushed = ioBus->read_responders.try_push_back({ this, &ProcessIOReadRequest }); RETURN_HR_IF(E_OUTOFMEMORY, !pushed);
		pushed = ioBus->write_responders.try_push_back({ this, &ProcessIOWriteRequest });
		if (!pushed)
		{
			ioBus->read_responders.remove([this](auto& responder) { return responder.Device == this; });
			return E_OUTOFMEMORY;
		}

		return S_OK;
	}

	~AYChip()
	{
		if (_ioBus)
		{
			_ioBus->read_responders.remove([this](auto& responder) { return responder.Device == this; });
			_ioBus->write_responders.remove([this](auto& responder) { return responder.Device == this; });
		}
		if (_sourceVoice)
			_sourceVoice->DestroyVoice();
	}

	HRESULT STDMETHODCALLTYPE Reset(SpectrumVariant variant) override
	{
		_time = 0;
		_variant = variant;
		memset(_registers, 0, sizeof(_registers));
		_selectedRegister = 0;
		memset(_toneCounters, 0, sizeof(_toneCounters));
		memset(_toneLevels, 0, sizeof(_toneLevels));
		_noiseCounter = 0;
		_noiseShiftRegister = 0x1FFFF;
		_noiseLevel = true;
		_envelopeCounter = 0;
		_envelopeCount = 15;
		_envelopeAttackMask = 0;
		_envelopeHolding = false;
		_sampleCount = 0;
		_samples[0] = audio_level_silence;
		SendSamplesToXAudio(_sourceVoice, _samples, 1);
		return S_OK;
	}

	BOOL STDMETHODCALLTYPE NeedSyncWithRealTime(UINT64* sync_time) override { return FALSE; }

	bool SimulateDeviceTo(UINT64 requested_time) override
	{
		WI_ASSERT(_time < requested_time);

		if (_variant != SpectrumVariant128)
		{
			_time = requested_time;
			return true;
		}

		while (_time < requested_time)
		{
			int sample = audio_level_silence;
			for (size_t channel = 0; channel < 3; channel++)
			{
				uint8_t volume = _registers[8 + channel] & 0x0F;
				if (_registers[8 + channel] & 0x10)
					volume = _envelopeCount ^ _envelopeAttackMask;
				int amplitude = volumeLevels[volume];
				bool toneEnabled = !(_registers[7] & (1 << channel));
				bool noiseEnabled = !(_registers[7] & (1 << (channel + 3)));
				bool channelHigh = (!toneEnabled || _toneLevels[channel]) && (!noiseEnabled || _noiseLevel);
				sample += channelHigh ? amplitude : -amplitude;
			}
			_samples[_sampleCount++] = static_cast<uint8_t>(std::clamp(sample, 0, 255));
			if (_sampleCount == _countof(_samples))
			{
				SendSamplesToXAudio(_sourceVoice, _samples, static_cast<uint32_t>(_sampleCount));
				_sampleCount = 0;
			}

			_time += audio_increment;
			for (size_t channel = 0; channel < 3; channel++)
			{
				uint16_t period = _registers[channel * 2] | ((_registers[channel * 2 + 1] & 0x0F) << 8);
				if (!period)
					period = 1;
				uint32_t togglePeriod = period * 16;
				_toneCounters[channel] += audio_increment;
				while (_toneCounters[channel] >= togglePeriod)
				{
					_toneCounters[channel] -= togglePeriod;
					_toneLevels[channel] = !_toneLevels[channel];
				}
			}

			uint32_t noisePeriod = _registers[6] & 0x1F;
			if (!noisePeriod)
				noisePeriod = 1;
			_noiseCounter += audio_increment;
			while (_noiseCounter >= noisePeriod * 32)
			{
				_noiseCounter -= noisePeriod * 32;
				uint32_t feedback = (_noiseShiftRegister ^ (_noiseShiftRegister >> 3)) & 1;
				_noiseShiftRegister = (_noiseShiftRegister >> 1) | (feedback << 16);
				_noiseLevel = (_noiseShiftRegister & 1) != 0;
			}

			uint16_t envelopePeriod = _registers[11] | (_registers[12] << 8);
			if (!envelopePeriod)
				envelopePeriod = 1;
			if (!_envelopeHolding)
			{
				_envelopeCounter += audio_increment;
				while (_envelopeCounter >= envelopePeriod * 512)
				{
					_envelopeCounter -= envelopePeriod * 512;
					if (_envelopeCount)
						--_envelopeCount;
					else
					{
						uint8_t shape = _registers[13];
						if (!(shape & 0x08))
							_envelopeHolding = true;
						else if (shape & 0x01)
						{
							if (shape & 0x02)
								_envelopeAttackMask ^= 0x0F;
							_envelopeHolding = true;
						}
						else
						{
							if (shape & 0x02)
								_envelopeAttackMask ^= 0x0F;
							_envelopeCount = 15;
						}
					}
				}
			}
		}

		return true;
	}

	static bool IsRegisterDataPort(WORD address) { return (address & 0xC002) == 0x8000; }

	static uint8_t ProcessIOReadRequest(IDevice* device, WORD address)
	{
		auto* chip = static_cast<AYChip*>(device);
		if (chip->_variant == SpectrumVariant128 && IsRegisterDataPort(address))
		{
			if (chip->_selectedRegister == 14)
				return (chip->_registers[7] & 0x40) ? 0xFF : chip->_registers[14];
			if (chip->_selectedRegister == 15)
				return 0xFF;
			return chip->_registers[chip->_selectedRegister];
		}
		return 0xFF;
	}

	static void ProcessIOWriteRequest(IDevice* device, WORD address, uint8_t value)
	{
		auto* chip = static_cast<AYChip*>(device);
		if (chip->_variant != SpectrumVariant128)
			return;
		if ((address & 0xC002) == 0xC000)
			chip->_selectedRegister = value & 0x0F;
		else if (IsRegisterDataPort(address))
		{
			static constexpr uint8_t registerMasks[16] = {
				0xFF, 0x0F, 0xFF, 0x0F, 0xFF, 0x0F, 0x1F, 0xFF,
				0x1F, 0x1F, 0x1F, 0xFF, 0xFF, 0x0F, 0xFF, 0xFF
			};
			chip->_registers[chip->_selectedRegister] = value & registerMasks[chip->_selectedRegister];
			if (chip->_selectedRegister == 13)
			{
				chip->_envelopeCounter = 0;
				chip->_envelopeCount = 15;
				chip->_envelopeAttackMask = (chip->_registers[13] & 0x04) ? 0x0F : 0;
				chip->_envelopeHolding = false;
			}
		}
	}
};

HRESULT STDMETHODCALLTYPE MakeAYChip(Bus* ioBus, IXAudio2* xaudio2, SpectrumVariant variant, wistd::unique_ptr<IDevice>& ppDevice)
{
	auto device = wil::make_unique_nothrow<AYChip>(); RETURN_IF_NULL_ALLOC(device);
	auto hr = device->InitInstance(ioBus, xaudio2, variant); RETURN_IF_FAILED(hr);
	ppDevice = std::move(device);
	return S_OK;
}
