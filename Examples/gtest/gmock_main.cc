// Copyright 2008, Google Inc.
// All rights reserved.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are
// met:
//
//     * Redistributions of source code must retain the above copyright
// notice, this list of conditions and the following disclaimer.
//     * Redistributions in binary form must reproduce the above
// copyright notice, this list of conditions and the following disclaimer
// in the documentation and/or other materials provided with the
// distribution.
//     * Neither the name of Google Inc. nor the names of its
// contributors may be used to endorse or promote products derived from
// this software without specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
// "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
// LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
// A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
// OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
// SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
// LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
// DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
// THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
// (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
// OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


#include <iostream>
#include <type_traits>

#include "gmock/gmock.h"
#include "gtest/gtest.h"

class SEHFailureListener : public ::testing::EmptyTestEventListener
{
	// Fired before the test case starts and before SetUpTestCase() is invoked.
    void OnTestCaseStart(const ::testing::TestCase& test_case) override
    {
		hasFailureCausedBySEH = false;
    }

    // Fired before each individual test starts: before the test fixture is constructed
    // and SetUp() is invoked.
    void OnTestStart(const ::testing::TestInfo& info) override
    {
		if (hasFailureCausedBySEH)
		{
			GTEST_SKIP() << "One of previous tests failed with SEH exception";
		}
    }

	void OnTestPartResult(const ::testing::TestPartResult& test_part_result) override
	{
		if (test_part_result.fatally_failed() && !hasFailureCausedBySEH)
		{
			const char seh_prefix[] = "SEH exception with code 0x";
			// https://stackoverflow.com/a/7142763
			// https://stackoverflow.com/a/41195467
			if (!strncmp(test_part_result.message(), seh_prefix, std::extent<decltype(seh_prefix)>::value - 1))
			{
				hasFailureCausedBySEH = true;
			}
		}
	}

	bool hasFailureCausedBySEH{false};
};

// class to turn ASSERT-failure into an exception
class ThrowListener : public testing::EmptyTestEventListener {
  void OnTestPartResult(const testing::TestPartResult& result) override {
    if (result.type() == testing::TestPartResult::kFatalFailure &&
        result.exception_type() == testing::TestPartResult::kGtestException) {
        throw testing::AssertionException(result);
    }
  }
};


// MS C++ compiler/linker has a bug on Windows (not on Windows CE), which
// causes a link error when _tmain is defined in a static library and UNICODE
// is enabled. For this reason instead of _tmain, main function is used on
// Windows. See the following link to track the current status of this bug:
// https://web.archive.org/web/20170912203238/connect.microsoft.com/VisualStudio/feedback/details/394464/wmain-link-error-in-the-static-library
// // NOLINT
#if GTEST_OS_WINDOWS_MOBILE
# include <tchar.h>  // NOLINT

GTEST_API_ int _tmain(int argc, TCHAR** argv) {
#else
GTEST_API_ int main(int argc, char** argv) {
#endif  // GTEST_OS_WINDOWS_MOBILE
  std::cout << "Running main() from gmock_main.cc\n";
  // Since Google Mock depends on Google Test, InitGoogleMock() is
  // also responsible for initializing Google Test.  Therefore there's
  // no need for calling testing::InitGoogleTest() separately.
  testing::InitGoogleMock(&argc, argv);

  auto& listeners = testing::UnitTest::GetInstance()->listeners();
  listeners.Append(new SEHFailureListener);
  listeners.Append(new ThrowListener);

  return RUN_ALL_TESTS();
}
