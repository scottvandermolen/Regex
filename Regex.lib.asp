<%
	' ASP Regular Expressions Library
	' 
	' Copyright (c) 2024, Scott Vander Molen; some rights reserved.
	' 
	' This work is licensed under a Creative Commons Attribution 4.0 International License.
	' To view a copy of this license, visit https://creativecommons.org/licenses/by/4.0/
	' 
	' @author  Scott Vander Molen
	' @version 2.0
	' @since   2024-01-02
	'
	' https://www.php.net/manual/en/ref.pcre.php

	' Returns a string or an array with pattern matches replaced, but only if matches were found.
	'
	' @param pattern Contains a regular expression indicating what to search for.
	' @param replacement A string which will replace the matched patterns. It may contain backreferences.
	' @param input A string or array of strings in which the replacements are being performed.
	' @return string An array of replaced strings if the input was an array, a string with replacements made if the input was a string, or null if the input was a string and no matches were found.
	function preg_filter(pattern, replacement, input)
		dim result
		
		if preg_match(pattern, input) = 1 then
			result = preg_replace(pattern, replacement, input)
		else
			result = null
		end if
		
		preg_filter = result
	end function

	' Returns an array consisting only of elements from the input array which matched the pattern.
	'
	' @param pattern Contains a regular expression indicating what to search for.
	' @param inputs An array of strings.
	' @return string An array containing strings that matched the provided pattern.
	function preg_grep(pattern, inputs)
		dim delimiter
		delimiter = ";;"
		
		' Check if the user specified case insensitivity.
		dim result
		dim input
		
		for each input in inputs
			' Only retain inputs that match the pattern.
			if preg_match(pattern, input) = 1 then
				' Add delimiter between multiple results.
				if result <> "" then
					result = result & delimiter
				end if
				
				result = result & input
			end if
		next
		
		preg_grep = Split(result, delimiter)
	end function

	' Finds the first match of a pattern in a string.
	'
	' @param pattern Contains a regular expression indicating what to search for.
	' @param input The string in which the search will be performed.
	' @return variant 1 if a match was found, 0 if no matches were found, or false if an error occurred.
	function preg_match(pattern, input)
		dim regEx
		set regEx = new RegExp
		
		' Check if pattern was delimited.
		if Left(pattern, 1) = "/" then
			dim flags 
			flags = Right(pattern, Len(pattern) - InStrRev(pattern, "/"))
			
			regex.Pattern = Mid(pattern, 2, InStrRev(pattern, "/") - 2)
			
			' Check if the user specified case insensitivity.
			if Instr(flags, "i") > 0 then
				regex.IgnoreCase = true
			end if
			
			' Check if the user specified multiline support.
			if Instr(flags, "m") > 0 then
				regex.Multline = true
			end if
		else
			regex.Pattern = pattern
		end if
		
		regEx.Global = False ' Only find the first occurrence of a match.
		
		dim result
		
		if regEx.test(input) then
			result = 1
		else
			result = 0
		end if
		
		preg_match = result
		set regEx = nothing
	end function

	' Finds all matches of a pattern in a string.
	'
	' @param pattern Contains a regular expression indicating what to search for.
	' @param input The string in which the search will be performed.
	' @return variant The number of matches found or false if an error occurred.
	function preg_match_all(pattern, input)
		dim regEx
		set regEx = new RegExp
		
		' Check if pattern was delimited.
		if Left(pattern, 1) = "/" then
			dim flags 
			flags = Right(pattern, Len(pattern) - InStrRev(pattern, "/"))
			
			regex.Pattern = Mid(pattern, 2, InStrRev(pattern, "/") - 2)
			
			' Check if the user specified case insensitivity.
			if Instr(flags, "i") > 0 then
				regex.IgnoreCase = true
			end if
			
			' Check if the user specified multiline support.
			if Instr(flags, "m") > 0 then
				regex.Multline = true
			end if
		else
			regex.Pattern = pattern
		end if
		
		regEx.Global = True ' Find all occurrences of a match.
		
		dim matches
		set matches = regEx.execute(input)
		
		preg_match_all = matches.count
		
		set regEx = nothing
		set matches = nothing
	end function

	' Returns a string where matches of a pattern (or an array of patterns) are replaced with a substring (or an array of substrings) in a given string.
	'
	' @param patterns Contains a regular expression or array of regular expressions.
	' @param replacements A replacement string or an array of replacement strings.
	' @param inputs The string or array of strings in which replacements are being performed.
	' @return string A string or an array of strings resulting from applying the replacements to the input string or strings.
	function preg_replace(patterns, replacements, inputs)
		dim regEx
		set regEx = new RegExp
		
		regEx.Global = True ' Find all occurrences of a match.
		
		dim delimiter
		delimiter = ";;"
		
		dim result
		
		' Check number of inputs.
		if isArray(inputs) then
			' Multiple inputs
			dim input
			
			for each input in inputs
				' Add delimiter between multiple results.
				if result <> "" then
					result = result & delimiter
				end if
				
				' Recursively call this function to process each input.
				result = result & preg_replace(patterns, replacements, input)
				
				' Split the string into an array.
				result = Split(result, delimiter)
			next
		else
			' Single input
			result = inputs
			
			' Check number of patterns.
			if isArray(patterns) then
				' Multiple patterns
				dim pattern
				
				for each pattern in patterns
					' Check number of replacements.
					if isArray(replacements) then
						' Multiple replacements
						dim loopCounter
						
						for loopCounter = 0 to UBound(patterns)
							' Check for equivalent replacement.
							if loopCounter > UBound(replacements) then
								' Use empty string if fewer replacements than patterns.
								result = preg_replace(pattern, "", result)
							else
								' Use equivalent replacement.
								result = preg_replace(pattern, replacements(loopCounter), result)
							end if
						next
					else
						' Single replacement
						result = preg_replace(pattern, replacements, result)
					end if
				next
			else
				' Single pattern, single replacement
				
				' Check if pattern was delimited.
				if Left(patterns, 1) = "/" then
					dim flags 
					flags = Right(patterns, Len(patterns) - InStrRev(patterns, "/"))
					
					regex.Pattern = Mid(patterns, 2, InStrRev(patterns, "/") - 2)
					
					' Check if the user specified case insensitivity.
					if Instr(flags, "i") > 0 then
						regex.IgnoreCase = true
					end if
					
					' Check if the user specified multiline support.
					if Instr(flags, "m") > 0 then
						regex.Multline = true
					end if
				else
					regex.Pattern = patterns
				end if
				
				result = regEx.replace(result, replacements)
			end if
		end if
		
		preg_replace = result
	end function
	
	' Given an expression and a callback, returns a string where all matches of the expression are replaced with the substring returned by the callback
	'
	' @param pattern A regular expression or array of regular expressions indicating what to search for.
	' @param callback A callback function which returns the replacement.
	' @param input The string or array of strings in which replacements are being performed.
	' @return string A string or an array of strings resulting from applying the replacements to the input string or strings.
	function preg_replace_callback(patterns, callback, inputs)
		dim regEx
		set regEx = new RegExp
		
		regEx.Global = True ' Find all occurrences of a match.
		
		dim delimiter
		delimiter = ";;"
		
		dim result
		
		' Check number of inputs.
		if isArray(inputs) then
			' Multiple inputs
			dim input
			
			for each input in inputs
				' Add a delimiter between multiple results.
				if result <> "" then
					result = result & delimiter
				end if
				
				' Recursively call this function to process each input.
				result = result & preg_replace_callback(patterns, callback, input)
				
				' Split the string into an array.
				result = Split(result, delimiter)
			next
		else
			' Single input
			result = inputs
			
			' Check number of patterns.
			if isArray(patterns) then
				' Multiple patterns
				dim pattern
				
				for each pattern in patterns
					result = preg_replace_callback(pattern, callback, result)
				next
			else
				' Single pattern
				
				' Check if pattern was delimited.
				if Left(patterns, 1) = "/" then
					dim flags 
					flags = Right(patterns, Len(patterns) - InStrRev(patterns, "/"))
					
					regex.Pattern = Mid(patterns, 2, InStrRev(patterns, "/") - 2)
					
					' Check if the user specified case insensitivity.
					if Instr(flags, "i") > 0 then
						regex.IgnoreCase = true
					end if
					
					' Check if the user specified multiline support.
					if Instr(flags, "m") > 0 then
						regex.Multline = true
					end if
				else
					regex.Pattern = patterns
				end if
				
				dim replacement
				set replacement = getRef(callback)
				result = regEx.replace(result, replacement)
			end if
		end if
		
		preg_replace_callback = result
	end function

	' Breaks a string into an array using matches of a regular expression as separators.
	'
	' @param pattern A regular expression determining what to use as a separator.
	' @param input The string that is being split.
	' @return string An array of substrings where each item corresponds to a part of the input string separated by a match of the regular expression.
	function preg_split(pattern, input)
		dim delimiter
		delimiter = ";;"
		
		dim result
		result = preg_replace(pattern, delimiter, input)
		preg_split = Split(result, delimiter)
	end function

	' Escapes characters that have a special meaning in regular expressions by putting a backslash in front of them.
	'
	' @param input The string to be escaped.
	' @param delimiter A single character indicating which delimiter the regular expression will use. Instances of this character in the input string will also be escaped with a backslash.
	' @return string A string or an array of strings resulting from applying the replacements to the input string or strings.
	function preg_quote(input, delimiter)
		dim result
		result = input
		
		' A list of special characters to be escaped.
		dim specialChars
		specialChars = Array("\", ".", "+", "*", "?", "[", "^", "]", "$", "(", ")", "{", "}", "=", "!", "<", ">", "|", ":", "-", "#")
		
		' A boolean for tracking whether the delimiter is included in the list of special characters.
		dim isSpecial
		isSpecial = false
		
		dim character
		
		for each character in specialChars
			' Check if the current special character is the delimiter.
			if delimiter = character then
				isSpecial = true
			end if
			
			result = Replace(result, character, "\" & character)
			
			' If the delimiter has not already been escaped, escape it now.
			if isSpecial = false then
				result = Replace(result, delimiter, "\" & delimiter)
			end if
		next
		
		preg_quote = result
	end function
%>