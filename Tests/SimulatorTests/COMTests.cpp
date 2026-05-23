
#include "CppUnitTest.h"
#include "shared/com.h"
#include "shared/inplace_function.h"
#include "wil/com.h"

using namespace Microsoft::VisualStudio::CppUnitTestFramework;

namespace Z80SimulatorTests
{
	TEST_CLASS(COMTests)
	{
		struct DECLSPEC_NOVTABLE DECLSPEC_UUID("AD5F0ADE-0FEF-452F-9294-CD3C21228400") ITestSink : IUnknown
		{
			virtual HRESULT STDMETHODCALLTYPE NotifyTest() = 0;
		};

		class TestSink : public ITestSink
		{
			ULONG _refCount = 0;
			stdext::inplace_function<HRESULT()> _callback;

		public:
			HRESULT InitInstance (stdext::inplace_function<HRESULT()> callback)
			{
				_callback = std::move(callback);
				return S_OK;
			}

			#pragma region IUnknown
			virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
			{
				if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<ITestSink>(this, riid, ppvObject))
					return S_OK;

				RETURN_HR(E_NOINTERFACE);
			}

			virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

			virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
			#pragma endregion

			virtual HRESULT STDMETHODCALLTYPE NotifyTest() override
			{
				return _callback();
			}
		};

		class TestSource : public IConnectionPointContainer
		{
			ULONG _refCount = 0;
			com_ptr<ConnectionPointImpl<ITestSink>> _cp;

		public:
			TestSource()
			{
				auto hr = MakeConnectionPoint(this, &_cp);
				Assert::AreEqual(S_OK, hr);
			}

			#pragma region IUnknown
			virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override
			{
				if (TryQI<IUnknown>(this, riid, ppvObject) || TryQI<IConnectionPointContainer>(this, riid, ppvObject))
					return S_OK;

				RETURN_HR(E_NOINTERFACE);
			}

			virtual ULONG STDMETHODCALLTYPE AddRef() override { return ++_refCount; }

			virtual ULONG STDMETHODCALLTYPE Release() override { return ReleaseST(this, _refCount); }
			#pragma endregion

			#pragma region IConnectionPointContainer
			virtual HRESULT STDMETHODCALLTYPE EnumConnectionPoints (IEnumConnectionPoints **ppEnum) override
			{
				Assert::Fail();
				return E_NOTIMPL;
			}

			virtual HRESULT STDMETHODCALLTYPE FindConnectionPoint (REFIID riid, IConnectionPoint **ppCP) override
			{
				if (riid == __uuidof(ITestSink))
					return copy_to(_cp.get(), ppCP);

				Assert::Fail();
				return E_NOTIMPL;
			}
			#pragma endregion

			void NotifyTestSinks()
			{
				_cp->Notify([](ITestSink* sink){ return sink->NotifyTest(); });
			}
		};

		TEST_METHOD(ConnectionPointUnadviseFromCallback)
		{
			HRESULT hr;

			auto source = wil::com_ptr_failfast<TestSource>(new TestSource());

			com_ptr<IConnectionPoint> cp;
			hr = source->FindConnectionPoint(__uuidof(ITestSink), &cp);
			Assert::AreEqual(S_OK, hr);

			DWORD cookie1;
			bool called1 = false;
			auto sink1 = wil::com_ptr_failfast(new TestSink());
			hr = sink1->InitInstance([&called1, &cp, &cookie1] { called1 = true; cp->Unadvise(cookie1); cookie1 = 0; return S_OK; });
			Assert::AreEqual(S_OK, hr);
			hr = cp->Advise(sink1, &cookie1);
			Assert::AreEqual(S_OK, hr);

			DWORD cookie2;
			bool called2 = false;
			auto sink2 = wil::com_ptr_failfast(new TestSink());
			hr = sink2->InitInstance([&called2, &cp, &cookie2] { called2 = true; cp->Unadvise(cookie2); cookie2 = 0; return S_OK; });
			Assert::AreEqual(S_OK, hr);
			hr = cp->Advise(sink2, &cookie2);
			Assert::AreEqual(S_OK, hr);

			DWORD cookie3;
			bool called3 = false;
			auto sink3 = wil::com_ptr_failfast(new TestSink());
			hr = sink3->InitInstance([&called3, &cp, &cookie3] { called3 = true; cp->Unadvise(cookie3); cookie3 = 0; return S_OK; });
			Assert::AreEqual(S_OK, hr);
			hr = cp->Advise(sink3, &cookie3);
			Assert::AreEqual(S_OK, hr);

			source->NotifyTestSinks();

			Assert::IsTrue(called1);
			Assert::IsTrue(called2);
			Assert::IsTrue(called3);
		}
	};
}

