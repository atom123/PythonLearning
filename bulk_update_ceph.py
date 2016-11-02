#!/usr/bin/python

####################################################################
# Author:       Jeffrey Guan
# Date:         2016-11-2
# Description:  Update ceph client.
#####################################################################

import os
import re


def cmd_install_cephclient(host_ip = ''):
  cmd = "ssh %s 'cd /home/upgrade-ceph-qemu; python \
		install-cephclient.py controller'"\
		% host_ip
  os.system(cmd)

def get_host_ips():
  ip_list = []

  # Open the file.
  file_obj = open('/etc/hosts')
  try:
    # Search the str "domain.tld".
    patt_save_str = re.compile(r'^.*domain.tld.*')

    # Read file content by lines.
    lines = file_obj.readlines()

    for line in lines:
      match = patt_save_str.search(line)
      if match:
        # Remove the "domain.tld" and tab from match.group(0).
        temp_str = match.group(0).strip("domain.tld").strip()
        temp_list  = temp_str.split("\t")

        # Get host ip.
        ip_list.append(temp_list[0])

  finally:
    file_obj.close()  

  return ip_list

if __name__ == "__main__":

  host_ips = []

  host_ips = get_host_ips()

  for ip in host_ips:
    cmd_install_cephclient(ip)
