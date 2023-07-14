#pragma once

#include <vector>
#include <mmeapi.h> // WAVEFORMATEX

// Converts unsigned char representation of 8, 16, 24 and 32 bit PCM data to signed float (ranging from -1.0 to 1.0)

inline bool PCMByteToFloat(WAVEFORMATEX &pFormat, std::vector<unsigned char> &vecIn, std::vector<float> &vecOut)
{
	vecOut.clear();

	if (pFormat.wFormatTag != WAVE_FORMAT_PCM)
		return false;

	size_t iAlign = (size_t)pFormat.nBlockAlign; // aka frame size, one frame = one sample of given bit depth for all channels
	size_t iAlignedSize = vecIn.size() / iAlign * iAlign;
	size_t iBitsPerSample = (size_t)pFormat.wBitsPerSample;
	size_t iBytesPerSample = iBitsPerSample / 8;
	float fMax = (float)pow(2U, iBitsPerSample - 1U);
	int32_t iTempData;

	if (iAlignedSize < iAlign || 
		(iBitsPerSample != 8 && iBitsPerSample != 16 && iBitsPerSample != 24 && iBitsPerSample != 32))
		return false;

	size_t iByte = 0;

	for (size_t i = 0; i < iAlignedSize; i += iBytesPerSample)
	{
		switch (iBitsPerSample)
		{
		case 8:
			vecOut.push_back(vecIn[i] / fMax - 1.0f);
			break;
		case 16:
			vecOut.push_back((int16_t)(vecIn[i] | (vecIn[i + 1] << 8)) / fMax);
			break;
		case 24:
			iTempData = vecIn[i] | (vecIn[i + 1] << 8) | (vecIn[i + 2] << 16);

			if (iTempData & (1 << 23))
				iTempData -= (1 << 24);

			vecOut.push_back(iTempData / fMax);
			break;
		case 32:
			vecOut.push_back((int32_t)(vecIn[i] | (vecIn[i + 1] << 8) | (vecIn[i + 2] << 16) | (vecIn[i + 3] << 24)) / fMax);
			break;
		}
	}

	return true;
}