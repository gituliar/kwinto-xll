// Taken from https://github.com/elnormous/HTTPRequest
#include "extern/HTTPRequest.hpp"

#include "framework/framework.h"

extern "C" __declspec(dllexport)
double
kwRPC(double x, double y)
{
	try
	{
		http::Request request{ "http://localhost:3000" };

		// send a get request
		const auto response = request.send("GET", "", {}, std::chrono::milliseconds{5000});
		//std::cout << std::string{ response.body.begin(), response.body.end() } << '\n'; // print the result
		return response.body.size();
	}
	catch (const std::exception& e)
	{
		//std::cerr << "Request failed, error: " << e.what() << '\n';
	}

	return -1;
}


extern "C" __declspec(dllexport)
int
xlAutoOpen(void)
{
	XLOPER12 xDLL;

	Excel12f(xlGetName, &xDLL, 0);

	Excel12f(xlfRegister, 0, 11, (LPXLOPER12)&xDLL,
		(LPXLOPER12)TempStr12(L"kwRPC"),
		(LPXLOPER12)TempStr12(L"BBB"),
		(LPXLOPER12)TempStr12(L"kwRPC"),
		(LPXLOPER12)TempStr12(L"x, y"),
		(LPXLOPER12)TempStr12(L"1"),
		(LPXLOPER12)TempStr12(L"myOwnCppFunctions"),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L""),
		(LPXLOPER12)TempStr12(L"Multiplies 2 numbers"),
		(LPXLOPER12)TempStr12(L""));


	// Free the XLL filename
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

	return 1;
}
