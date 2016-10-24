#!/usr/bin/python

####################################################################
# Author:      Jeffrey Guan
# Date:        2016-10-24
# Description: set the environment parameters in python script.
#              For example, if we want to run "nova list" in python 
#              script before "source openrc" in CLI, we can safely 
#              use this script. 
#
#              A part of "/root/openrc" is:
#               #!/bin/sh
#               export LC_ALL=C
#               export OS_NO_CACHE='true'
#               export OS_TENANT_NAME='admin'
#               export OS_PROJECT_NAME='admin'
#               export OS_USERNAME='admin'
#               export OS_PASSWORD='admin'
#####################################################################

import os
import re

# Open the file.
file_obj = open('/root/openrc')

try:
  # Search the str started with "export" and contains "=".
  patt_save_str = re.compile(r'^export.*=.*')
  # Search "=".
  patt_rm_str = re.compile(r'=')

  # Read file content by lines.
  lines = file_obj.readlines()

  for line in lines:
    match = patt_save_str.search(line)
    if match:
      # Remove the "export" and "" from match.group(0).
      temp_str = match.group(0).strip("export").strip()

      # Split the str into a list by "=".
      environ_value_dic = patt_rm_str.split(temp_str)

      # Set the value for each env parameters.
      os.environ[environ_value_dic[0]] = environ_value_dic[1].strip("'")

  # Print the env and values. Or we can run "env" on the CLI to check 
  # all env and values.
  #
  #print os.getenv('OS_PROJECT_NAME')

  # Test nova command.
  os.system('nova list')

finally:
  file_obj.close()
