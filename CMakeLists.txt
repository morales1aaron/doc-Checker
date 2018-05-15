cmake_minimum_required(VERSION 3.5)
project(doc_Checker CXX)

# require a C++11 compiler for all targets
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(LIB_SOURCE
  doc_Checker.cpp doc_Checker.hpp)

add_executable(checker main.cpp ${LIB_SOURCE})

