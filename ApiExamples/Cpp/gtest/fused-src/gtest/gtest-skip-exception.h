#pragma once
// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
// 29/03/2019 by Tengiz Sharafiev <tengiz.sharafiev@aspose.com>

#include <stdexcept>

namespace testing {

// Exception which will mark test case as skipped
class SkipException : public ::std::runtime_error {
 public:
  SkipException() : runtime_error("skipped") {}
};

}
