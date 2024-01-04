<%@ CodePage=65001 Language="VBScript"%>
<% Option Explicit %>
<!--#include file="Regex.lib.asp"-->
<%
	' ASP Regular Expressions Unit Tests
	' 
	' Copyright (c) 2024, Scott Vander Molen; some rights reserved.
	' 
	' This work is licensed under a Creative Commons Attribution 4.0 International License.
	' To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/
	' 
	' @author  Scott Vander Molen
	' @version 2.0
	' @since   2024-01-03
	'
	' https://www.php.net/manual/en/ref.pcre.php

	' Ensure that UTF-8 encoding is used instead of Windows-1252
	Session.CodePage = 65001
	Response.CodePage = 65001
	Response.CharSet = "UTF-8"
	Response.ContentType = "text/html"
	
	' Framework for running tests and building result strings.
	' 
	' @param input The specified value.
	' @param expected The expected value.
	' @param actual The actual value.
	' @return string The results of the test.
	function testFramework(input, expected, actual)
		dim result
		dim resultText
		
		if actual = expected or (isnull(actual) and isnull(expected)) then
			result = true
			resultText = "successful"
		else
			result = false
			resultText = "failed"
		end if
		
		dim returnString
		returnString = "Input:     " & input & vbCrLf & _
			"Expected:  " & expected & vbCrLf & _
			"Actual:    " & actual & vbCrLf & _
			"Result:    Test " & resultText &  "!" & vbCrLf & vbCrLf
		
		testFramework = returnString
	end function
	
	' Create an HTML container for our output.
	Response.Write "<!DOCTYPE html>" & vbCrLf
	Response.Write "<html lang=""en"">" & vbCrLf
	Response.Write "<meta http-equiv=""Content-Type"" content=""text/html;charset=UTF-8"" />" & vbCrLf
	Response.Write "<body>" & vbCrLf
	
	' Display code header
	Response.Write "<pre>"
	Response.Write "/***************************************************************************************\" & vbCrLf
	Response.Write "| ASP Regular Expressions Unit Tests                                                    |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "| Copyright (c) 2024, Scott Vander Molen; some rights reserved.                         |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "| This work is licensed under a Creative Commons Attribution 4.0 International License. |" & vbCrLf
	Response.Write "| To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/    |" & vbCrLf
	Response.Write "|                                                                                       |" & vbCrLf
	Response.Write "\***************************************************************************************/" & vbCrLf
	Response.Write "</pre>"
	
	' Run unit tests
	Response.Write "<pre>"
	
	dim input
	dim expected
	dim actual
	
	input = "The quick brown fox jumps over the lazy dog. /brown/ -> red"
	expected = "The quick red fox jumps over the lazy dog."
	actual = preg_filter("/brown/", "red", "The quick brown fox jumps over the lazy dog.")
	Response.Write "Unit Test: preg_filter()" & vbCrLf & testFramework(input, expected, actual)

	input = "The quick brown fox jumps over the lazy dog. /black/ -> red"
	expected = null
	actual = preg_filter("/black/", "red", "The quick brown fox jumps over the lazy dog.")
	Response.Write "Unit Test: preg_filter()" & vbCrLf & testFramework(input, expected, actual)

	input = "black, blue, brown, red"
	expected = "black, blue, brown"
	actual = Join(preg_grep("/^b/", Array("black", "blue", "brown", "red")), ", ")
	Response.Write "Unit Test: preg_grep()" & vbCrLf & testFramework(input, expected, actual)

	input = "The quick brown fox jumps over the lazy dog. /dog/"
	expected = 1
	actual = preg_match("/dog/", "The quick brown fox jumps over the lazy dog.")
	Response.Write "Unit Test: preg_match()" & vbCrLf & testFramework(input, expected, actual)

	input = "The quick brown fox jumps over the lazy dog. /cat/"
	expected = 0
	actual = preg_match("/cat/", "The quick brown fox jumps over the lazy dog.")
	Response.Write "Unit Test: preg_match()" & vbCrLf & testFramework(input, expected, actual)

	input = "The rain in Spain falls mainly on the plain. /the/i"
	expected = 2
	actual = preg_match_all("/the/i", "The rain in Spain falls mainly on the plain.")
	Response.Write "Unit Test: preg_match_all()" & vbCrLf & testFramework(input, expected, actual)

	function CallbackTest(singleMatch, position, fullString)
		CallbackTest = UCase(Left(singleMatch, 1)) & LCase(Mid(singleMatch, 2))
	end function

	input = "The quick brown fox jumps over the lazy dog. /brown/ -> red"
	expected = "The quick red fox jumps over the lazy dog."
	actual = preg_replace("/brown/", "red", "The quick brown fox jumps over the lazy dog.")
	Response.Write "Unit Test: preg_replace()" & vbCrLf & testFramework(input, expected, actual)

	input = "The quick brown fox jumps over the lazy dog. /\b\w+\b/"
	expected = "The Quick Brown Fox Jumps Over The Lazy Dog."
	actual = preg_replace_callback("/\b\w+\b/", "CallbackTest", "The quick brown fox jumps over the lazy dog.")
	Response.Write "Unit Test: preg_replace_callback()" & vbCrLf & testFramework(input, expected, actual)

	input = "The quick brown fox jumps over the lazy dog. /\s/"
	expected = "The, quick, brown, fox, jumps, over, the, lazy, dog."
	actual = Join(preg_split("/\s/", "The quick brown fox jumps over the lazy dog."), ", ")
	Response.Write "Unit Test: preg_split()" & vbCrLf & testFramework(input, expected, actual)

	input = "I'll give you $5 for it."
	expected = "I'll give you \$5 for it\."
	actual = preg_quote("I'll give you $5 for it.", "/")
	Response.Write "Unit Test: preg_quote()" & vbCrLf & testFramework(input, expected, actual)

	Response.Write "</pre>" & vbCrLf

	' Close the HTML container.
	Response.Write "</body>" & vbCrLf
	Response.Write "</html>"
%>