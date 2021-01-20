import base64
import json
import logging
import os
import time
from datetime import datetime
from datetime import timedelta

import msal
import pytz
import requests

# Optional logging
# logging.basicConfig(level=logging.DEBUG)

config = json.load(open("gpt_parameters.json"))

# Create a preferably long-lived app instance which maintains a token cache.
app = msal.ConfidentialClientApplication(
    config["client_id"], authority=config["authority"],
    client_credential=config["secret"],
    # token_cache=...  # Default cache is in memory only.
    # You can learn how to use SerializableTokenCache from
    # https://msal-python.rtfd.io/en/latest/#msal.SerializableTokenCache
)

# The pattern to acquire a token looks like this.
result = None

# Firstly, looks up a token from cache
# Since we are looking for token for the current app, NOT for an end user,
# notice we give account parameter as None.
result = app.acquire_token_silent(config["scope"], account=None)

if not result:
    logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
    result = app.acquire_token_for_client(scopes=config["scope"])

from_time, to_time = datetime.now() - timedelta(minutes=100), datetime.now()
from_time = from_time.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
to_time = to_time.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
flag = 0
while 1:
    if flag == 0:
        from_, to_ = from_time, to_time
        # print(from_time, to_time)
    print(from_, to_)
    flag = 1
    ##all code here
    if "access_token" in result:
        flag = 0
        while 1:
            if flag == 0:
                query1 = f"https://graph.microsoft.com/v1.0/users/ilsmediclaim@gptgroup.co.in" \
                         f"/mailFolders/inbox/messages?$filter=(receivedDateTime ge {from_}) " \
                         f"and (receivedDateTime le {to_})"
            flag = 1
            graph_data2 = requests.get(query1,
                                       headers={'Authorization': 'Bearer ' + result['access_token']}, ).json()
            for i in graph_data2['value']:
                format = "%Y-%m-%dT%H:%M:%SZ"
                b = datetime.strptime(i['receivedDateTime'], format).replace(tzinfo=pytz.utc).astimezone(
                    pytz.timezone('Asia/Kolkata')).replace(
                    tzinfo=None)
                print(i['receivedDateTime'], b, i['subject'])
                print(i['sender']['emailAddress']['address'])
                if 'hasAttachments' in i:
                    q = f"https://graph.microsoft.com/v1.0/users/ilsmediclaim@gptgroup.co.in/mailFolders/inbox/messages/{i['id']}/attachments"
                    attach_data = requests.get(q,
                                               headers={'Authorization': 'Bearer ' + result['access_token']}, ).json()
                    for j in attach_data['value']:
                        if '@odata.mediaContentType' in j:
                            print(j['@odata.mediaContentType'], j['name'])
                            with open(os.path.join('new_attach', j['name']), 'w+b') as fp:
                                fp.write(base64.b64decode(j['contentBytes']))
                                print('wrote', j['name'])

    ##
    time.sleep(5)
    now_time = datetime.now().astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    # print(to_time, now_time)
    from_ = to_
    to_ = now_time
pass
