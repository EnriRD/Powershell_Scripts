import requests
# Use self signed certs
requests.packages.urllib3.disable_warnings()
# Credentials
USER = 'developer'
PASS = 'C1sco12345'
url = 'https://ios-xe-mgmt.cisco.com/restconf/data/ietf-interfaces:interfaces'

headers = {
'Accept': 'application/yang-data+json',
'Content-Type': 'application/yang-data+json'
}

int_number = 5

for x in range(int_number):
        ipaddr = '1.2.3.' + str(x)
        print('Creating loopback :' + ipaddr)
# Important – make sure your spacing for the
# subsequent code is right – add 4 spaces to
# left of the code so it is part of the loop.
        payload = '\
                {\
                        "ietf-interfaces:interface": {\
                                "name": "Loopback123' + str(x) + '",\
                                "description": "Added with RESTCONF",\
                                "type": "iana-if-type:softwareLoopback",\
                                "enabled": true,\
                                "ietf-ip:ipv4": {\
                                        "address": [\
                                                {\
                                                        "ip": "1.2.3.' + str(x) + '",\
                                                        "netmask": "255.255.255.255"\
                                                }\
                                        ]\
                                }\
                        }\
                }'
#print (payload)
        response = requests.request('POST',url, auth=(USER, PASS),
                headers=headers, data = payload, verify=False)
        print('Status Code:' + str(response.status_code))
        print('Response Text:' + response.text)
