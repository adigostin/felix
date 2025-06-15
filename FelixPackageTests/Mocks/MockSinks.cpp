
#include "pch.h"
#include "../Mocks.h"
#include "shared/vector_nothrow.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	struct MockPropertyNotifySink : IPropertyNotifySink, IMockPropertyNotifySink
	{
		ULONG _refCount = 0;
		vector_nothrow<DISPID> _changed;

		#pragma region IUnknown
		virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
		{
			if (   TryQI<IUnknown>(static_cast<IPropertyNotifySink*>(this), riid, ppvObject)
				|| TryQI<IPropertyNotifySink>(this, riid, ppvObject)
				|| TryQI<IMockPropertyNotifySink>(this, riid, ppvObject)
			)
				return S_OK;

			Assert::Fail();
		}

		virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

		virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
		#pragma endregion

		#pragma region IPropertyNotifySink
		virtual HRESULT STDMETHODCALLTYPE OnChanged (DISPID dispID) override
		{
			auto it = _changed.find(dispID);
			if (it == _changed.end())
				_changed.try_push_back(dispID);
			return S_OK;
		}

		virtual HRESULT STDMETHODCALLTYPE OnRequestEdit (DISPID dispID) override
		{
			return E_NOTIMPL;
		}
		#pragma endregion

		#pragma region IMockPropertyNotifySink
		virtual bool IsChanged (DISPID dispid) const override
		{
			return _changed.find(dispid) != _changed.end();
		}
		#pragma endregion
	};

	com_ptr<IMockPropertyNotifySink> MakeMockPropertyNotifySink()
	{
		return new MockPropertyNotifySink();
	};
}
