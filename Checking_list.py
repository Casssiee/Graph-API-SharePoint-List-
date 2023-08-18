import sharepoint 
import json
from dotenv import load_dotenv
import os
load_dotenv()

sharepoint_connection = sharepoint.MicrosoftGraph(client_id=os.getenv('app_registration_client_id'),
                                        tenant_id=os.getenv('app_registration_tenant_id'),
                                        client_credential=os.getenv('app_registration_client_secret'),
                                        username='the username',
                                        password=os.getenv('sharepoint_password'),
                                        scope=["Sites.ReadWrite.All"])

#print(sharepoint_connection)
sharepoint_response = sharepoint_connection.get_sharepoint_list(tenant_name='the tenant name',
                                                                team_id='the team id',
                                                                list_id='the list id')
                                                
print(sharepoint_response.json())
