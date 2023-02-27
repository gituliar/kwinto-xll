#pragma once

// Taken from https://github.com/elnormous/HTTPRequest
#include "extern/HTTPRequest.hpp"
#include "framework/framework.h"
#include "nlohmann/json.hpp"

using json = nlohmann::json;

using Error = std::string;

using namespace std::string_literals;