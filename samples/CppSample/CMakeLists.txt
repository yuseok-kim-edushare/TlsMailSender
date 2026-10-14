cmake_minimum_required(VERSION 3.20)
project(TlsMailSenderCppSample CXX)

add_executable(TlsMailSenderCppSample main.cpp)

if (MSVC)
    target_compile_definitions(TlsMailSenderCppSample PRIVATE UNICODE _UNICODE)
    target_compile_options(TlsMailSenderCppSample PRIVATE /W4 /utf-8)
endif()
