"""Process DMARC report mail items from an MS365 online mail folder via the GRAPH API using app_id and secret.

fill config template below and save this in `config.cfg` (same folder)

[graph-dmarc-mail]
clientId = xxxxx
clientSecret = xxxxx
tenantId = xxxxx
mailbox = something@example.com
mailbox_folder = Inbox
"""

import base64
import gzip
import io
from typing import Any
from zipfile import ZipFile

import configparser
from datetime import datetime, timedelta

import json
import requests
from configparser import SectionProxy

GRAPH_URL = "https://graph.microsoft.com{}"
LOGIN_URL = "https://login.microsoftonline.com/{}/oauth2/v2.0/token"

class Graph:
    settings: SectionProxy

    def __init__(self, config: SectionProxy) -> None:
        self._token: str | None = None
        self._token_expires: datetime | None = None
        self._folder_id: str | None = None
        self.settings = config
        self.client_id: str = self.settings['clientId']
        self.client_secret: str = self.settings['clientSecret']
        self.tenant_id: str = self.settings['tenantId']
        self.mailbox: str = self.settings['mailbox']
        self.mailbox_folder: str = self.settings['mailbox_folder']


    def get_token(self) -> None:
        """Get an GRAPH API token."""
        
        def _get_fresh_token():
            endpoint = LOGIN_URL.format(self.tenant_id)
            data = {
                "grant_type": "client_credentials",
                "client_id" : self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default"
            }
            response = requests.post(endpoint, data=data)
            if response.status_code == 200:
                token_data = json.loads(response.text)
                self._token = token_data["access_token"]
                self._token_expires = datetime.now() + timedelta(seconds=token_data["expires_in"])
                return
            self._token = None

        if self._token is None or self._token_expires + timedelta(seconds=300) > datetime.now():
            _get_fresh_token()

    def api_get_request(self, endpoint) -> dict | None:
        """Perform API get request."""
        if not self._token:
            return None
        headers = { "Authorization": f"Bearer {self._token}" }
        response = requests.get(endpoint, headers=headers)
        if response.status_code != 200:
            return None
        return json.loads(response.text)["value"]

    def api_delete_request(self, endpoint) -> bool:
        """Perform API delete request."""
        if not self._token:
            return False
        headers: dict[str, Any] = { "Authorization": f"Bearer {self._token}" }
        response = requests.delete(endpoint, headers=headers)
        if response.status_code != 204:
            return False
        return True

    def get_folder_id(self) -> None:
        """Get the folder id."""
        folder_id = None
        self.get_token()
        if not self._token or self._folder_id:
            return
        endpoint = GRAPH_URL.format(f"/v1.0/users/{self.mailbox}/mailFolders")
        for folder in self.api_get_request(endpoint):
            if folder["displayName"] == self.mailbox_folder:
                folder_id = folder["id"]
                continue
        self._folder_id = folder_id
        
    def process_dmarc_mail_items(self) -> dict[str, bytes] :
        """Demonstrate fetching XML from dmarc report mails."""
        
        processed_items: dict[str, bytes] = {}

        def get_extention(name):
            if name is None:
                return None
            parts = name.split('.')
            return parts[len(parts) - 1]
        
        def get_dmarc_xml(attachment):
            ext = get_extention(attachment['name'])
            content_type = attachment['contentType']
            raw_data = base64.b64decode(attachment['contentBytes'])
            if content_type == "application/gzip" or ext == "gz" or ext == "gzip":
                try:
                    return gzip.decompress(raw_data)
                except Exception:
                    return None
            elif content_type == "application/zip" or ext == "zip":
                try:
                    with io.BytesIO(raw_data) as zip:
                        zipfile = ZipFile(zip) 
                        if files := zipfile.filelist:
                            file_name = files[0].filename
                            print(f"File name: {file_name}")
                            if get_extention(file_name) == "xml":
                                return zipfile.read(file_name)
                    return None
                except Exception:
                    return None
                
            return None
                

        self.get_folder_id()
        endpoint: str = GRAPH_URL.format(f"/v1.0/users/{self.mailbox}/mailFolders/{self._folder_id}/messages")
        for item in self.api_get_request(endpoint):
            print(f"Received: {datetime.strptime(item['receivedDateTime'], '%Y-%m-%dT%H:%M:%SZ')}")
            print(f"Subject: {item['subject']}")
            attachments_endpoint = GRAPH_URL.format(f"/v1.0/users/{self.mailbox}/messages/{item['id']}/attachments")
            
            # Take first attachment
            if (attachments := self.api_get_request(attachments_endpoint)) and (attachment := attachments[0]):
                print(f"Name: {attachment['name']}")
                # application/gzip, application/octet-stream + ext = .gz
                # application/zip, application/octet-stream + ext = .zip
                print(f"Content type: {attachment['contentType']}")
                xml_data = get_dmarc_xml(attachment)
                if xml_data:
                    processed_items[item['id']] = xml_data
                    print("Attachment data base64:")
                    print(xml_data)
                    

        return processed_items

    def delete_mail_items(self, items: list[str]) -> None:
        """Demonstrate deleting the processed mail items."""
        id: str
        for id in items:
            endpoint = GRAPH_URL.format(f"/v1.0/users/{self.mailbox}/messages/{id}")
            if not self.api_delete_request(endpoint):
                print("Unable to delete processed item! {id}")


def main() -> None:
    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg'])
    settings = config['graph-dmarc-mail']

    graph: Graph = Graph(settings)
    
    items: dict[str, bytes] = graph.process_dmarc_mail_items().items()
    for id, xml_data in items:
        print(f"ID: {id}: Bytes: {len(xml_data)}")
        
    # delete processed mail items
    graph.delete_mail_items(dict(items).keys())
    
if __name__ == '__main__':
    main()