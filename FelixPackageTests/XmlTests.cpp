
#include "pch.h"
#include "Mocks.h"
#include "Z80Xml.h"
#include "FelixPackage.h"
#include "XmlTests_h.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace FelixTests
{
	TEST_CLASS(XmlTests)
	{
		TEST_METHOD(PropertyWithoutSetterNotSaved)
		{
			struct TestClass : IProperties1
			{
				ULONG _refCount = 0;

				#pragma region IUnknown
				virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
				{
					return (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<IProperties1>(this, riid, ppvObject)) ? S_OK : E_NOINTERFACE;
				}

				virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

				virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
				#pragma endregion

				IMPLEMENT_IDISPATCH(IID_IProperties1)

				#pragma region IProperties1
				virtual HRESULT STDMETHODCALLTYPE get_OutputFilename (BSTR *pbstrOutputFilename) override
				{
					Assert::Fail(L"Should not be called since there's no corresponding setter, so no way to restore the property from XML.");
				}
				#pragma endregion
			};

			auto tc = wil::com_ptr_failfast(new TestClass());
			IStream* raw = SHCreateMemStream (nullptr, 0); Assert::IsNotNull(raw);
			com_ptr<IStream> stream;
			stream.attach(raw);
			auto hr = SaveToXml (tc, L"Temp", 0, stream);
			Assert::IsTrue(SUCCEEDED(hr));
		}

		TEST_METHOD(PropertyWithDefaultValue)
		{
			static const wchar_t DefaultValue[] = L"Def";

			struct TestClass : IProperties2, IVsPerPropertyBrowsing
			{
				ULONG _refCount = 0;
				wil::unique_process_heap_string _value = wil::make_process_heap_string_failfast(DefaultValue);
				bool _getterCalled = false;

				#pragma region IUnknown
				virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
				{
					return (TryQI<IUnknown>(static_cast<IProperties2*>(this), riid, ppvObject)
						|| TryQI<IProperties2>(this, riid, ppvObject)
						|| TryQI<IVsPerPropertyBrowsing>(this, riid, ppvObject)) ? S_OK : E_NOINTERFACE;
				}

				virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

				virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
				#pragma endregion

				IMPLEMENT_IDISPATCH(IID_IProperties2);

				#pragma region IProperties2
				virtual HRESULT STDMETHODCALLTYPE get_OutputFilename (BSTR* pbstrOutputFilename) override
				{
					_getterCalled = true;
					*pbstrOutputFilename = FAIL_FAST_IF_NULL_ALLOC(SysAllocString(_value.get()));
					return S_OK;
				}

				virtual HRESULT STDMETHODCALLTYPE put_OutputFilename (BSTR bstrOutputFilename) override { Assert::Fail(); }
				#pragma endregion

				#pragma region IVsPerPropertyBrowsing
				virtual HRESULT STDMETHODCALLTYPE HideProperty (DISPID dispid, BOOL *pfHide) override { Assert::Fail(); }

				virtual HRESULT STDMETHODCALLTYPE DisplayChildProperties (DISPID dispid, BOOL *pfDisplay) override { Assert::Fail(); }

				virtual HRESULT STDMETHODCALLTYPE GetLocalizedPropertyInfo (DISPID dispid, LCID localeID, BSTR *pbstrLocalizedName, BSTR *pbstrLocalizeDescription) override { Assert::Fail(); }

				virtual HRESULT STDMETHODCALLTYPE HasDefaultValue (DISPID dispid, BOOL *fDefault) override
				{
					*fDefault = !wcscmp(_value.get(), DefaultValue);
					return S_OK;
				}

				virtual HRESULT STDMETHODCALLTYPE IsPropertyReadOnly (DISPID dispid, BOOL *fReadOnly) override { Assert::Fail(); }

				virtual HRESULT STDMETHODCALLTYPE GetClassName (BSTR *pbstrClassName) override { Assert::Fail(); }

				virtual HRESULT STDMETHODCALLTYPE CanResetPropertyValue (DISPID dispid, BOOL *pfCanReset) override { Assert::Fail(); }

				virtual HRESULT STDMETHODCALLTYPE ResetPropertyValue (DISPID dispid) override { Assert::Fail(); }
				#pragma endregion
			};

			auto tc = wil::com_ptr_failfast(new TestClass());
			IStream* raw = SHCreateMemStream (nullptr, 0); Assert::IsNotNull(raw);
			com_ptr<IStream> stream;
			stream.attach(raw);

			auto hr = SaveToXml (tc, L"Temp", 0, stream);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsFalse(tc->_getterCalled, L"Should not be called since we said it has a default value (in the implementation of HasDefaultValue()).");

			hr = SaveToXml (tc, L"Temp", SAVE_XML_FORCE_SERIALIZE_DEFAULTS, stream);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(tc->_getterCalled, L"Should be called because we passed SAVE_XML_FORCE_SERIALIZE_DEFAULTS.");
		}

		TEST_METHOD(PropertyWithDefaultValueSaved)
		{
			// Even if the property has a default value, without IVsPerPropertyBrowsing implemented there's no way to recognize that, so it gets saved.
			static const wchar_t DefaultValue[] = L"Def";

			struct TestClass : IProperties2
			{
				ULONG _refCount = 0;
				wil::unique_process_heap_string _value = wil::make_process_heap_string_failfast(DefaultValue);
				bool _getterCalled = false;

				#pragma region IUnknown
				virtual HRESULT STDMETHODCALLTYPE QueryInterface (REFIID riid, void** ppvObject) override
				{
					return (TryQI<IUnknown>(static_cast<IProperties2*>(this), riid, ppvObject)
						|| TryQI<IProperties2>(this, riid, ppvObject)) ? S_OK : E_NOINTERFACE;
				}

				virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

				virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
				#pragma endregion

				IMPLEMENT_IDISPATCH(IID_IProperties2);

				#pragma region IProperties2
				virtual HRESULT STDMETHODCALLTYPE get_OutputFilename (BSTR *pbstrOutputFilename) override
				{
					_getterCalled = true;
					*pbstrOutputFilename = FAIL_FAST_IF_NULL_ALLOC(SysAllocString(_value.get()));
					return S_OK;
				}

				virtual HRESULT STDMETHODCALLTYPE put_OutputFilename (BSTR bstrOutputFilename) override { Assert::Fail(); }
				#pragma endregion
			};

			auto tc = wil::com_ptr_failfast(new TestClass());
			IStream* raw = SHCreateMemStream (nullptr, 0); Assert::IsNotNull(raw);
			com_ptr<IStream> stream;
			stream.attach(raw);
			auto hr = SaveToXml (tc, L"Temp", 0, stream);
			Assert::IsTrue(SUCCEEDED(hr));
			Assert::IsTrue(tc->_getterCalled);
		}
	};
}
