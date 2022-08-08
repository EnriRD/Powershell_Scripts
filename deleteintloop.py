import requests

# Use self signed certs
requests.packages.urllib3.disable_warnings()

# Credentials
USER = 'developer'
PASS = 'C1sco12345'

payload = {}

headers = {
'Accept': 'application/yang-data+json',
}

int_number = 5

for x in range(int_number):
        intname = 'Loopback123' + str(x)
        print('Deleting ' + intname)
        url = "https://ios-xe-mgmt.cisco.com/restconf/data/ietf-interfaces:interfaces/interface=" + intname
        response = requests.request("DELETE",url, auth=(USER, PASS),headers = headers, data = payload, verify=False)
        print('Status Code:' + str(response.status_code))
        print('Response Text:' + response.text)
