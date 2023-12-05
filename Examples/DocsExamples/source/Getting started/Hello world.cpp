#include "Hello world.h"

using namespace Aspose::Words;
namespace DocsExamples {
    namespace Programming_with_Documents {
        namespace gtest_test {

            class HelloWorld : public ::testing::Test
            {
            protected:
                static System::SharedPtr<::DocsExamples::Getting_started::HelloWorld> s_instance;

                void SetUp() override
                {
                    s_instance->SetUp();
                };

                static void SetUpTestCase()
                {
                    s_instance = System::MakeObject<::DocsExamples::Getting_started::HelloWorld>();
                    s_instance->OneTimeSetUp();
                };

                static void TearDownTestCase()
                {
                    s_instance->OneTimeTearDown();
                    s_instance = nullptr;
                };
            };

            System::SharedPtr<::DocsExamples::Getting_started::HelloWorld> HelloWorld::s_instance;

            TEST_F(HelloWorld, SimpleHelloWorld)
            {
                s_instance->SimpleHelloWorld();
            }
        }
    }
} // namespace DocsExamples::Getting_started::gtest_test
