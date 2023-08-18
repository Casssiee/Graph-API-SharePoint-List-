import sharepoint 
import json
from dotenv import load_dotenv
import os
load_dotenv()

sharepoint_connection = sharepoint.MicrosoftGraph(client_id=os.getenv('app_registration_client_id'),
                                        tenant_id=os.getenv('app_registration_tenant_id'),
                                        client_credential=os.getenv('app_registration_client_secret'),
                                        username='blueprism.bot@sydney.edu.au',
                                        password=os.getenv('sharepoint_password'),
                                        scope=["Sites.ReadWrite.All"])

#print(sharepoint_connection)
sharepoint_response = sharepoint_connection.get_sharepoint_list(tenant_name='unisyd',
                                                                team_id='AIHub',
                                                                list_id='c0135071-12ff-4186-b9f8-d3939d7c54a4')
                                                
print(sharepoint_response.json())