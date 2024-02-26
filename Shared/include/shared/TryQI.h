#pragma once

template<typename ITo>
static bool TryQI (ITo* from, REFIID riid, void** ppvObject)
{
	if (__uuidof(from) == riid)
	{
		*ppvObject = from;
		from->AddRef();
		return true;
	}
	return false;
}
