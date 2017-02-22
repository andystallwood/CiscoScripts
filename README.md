# CiscoConfigScript

Apply Configs to a number of Cisco devices via SecureCRT

Input File needs to be in the same format as Commands.csv. Be careful when copy and pasting into the CSV input files. - can be - or it can be equivalent to -- if pasting from Word, Outlook, etc.

Category - Is this a "config" command to enter or a show command to "test"

CommandStart - Start of the command line to enter onto the Cisco device

Parameter - Parameter to insert between CommandStart and CommandEnd. NA if no paramater

CommandEnd - End of command line to enter after parameter. Leave blank if nothing to add after parameter. DON'T put NA

Note: The command that gets sent to the device will be a Concatenation of CommandStart, Parameter and CommandEnd (spaces are automatically added by the script, so they don't need adding to the input file)

Prompt After Command - What Prompt should the Cisco device provide back after completing the command (NA for a test)

Expected Response - What text should appear int he response to a test show command (NA for a config)

WarnorFail - If the test fails, should the script continue onto the next line, or exit that device (NA for a config)
