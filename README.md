# CiscoConfigScript

Apply Configs to a number of Cisco devices via SecureCRT

Input File needs to be in the same format as Commands.csv

Category - Is this a "config" command to enter or a show command to "test"

Command - What command to enter onto the Cisco device

Prompt After Command - What Prompt should the Cisco device provide back after completing the command (NA for a test)

Expected Response - What text should appear int he response to a test show command (NA for a config)

Warn or Fail - If the test fails, should the script continue onto the next line, or exit that device (NA for a config)
