'''
pip install paramiko
pip instell msal
pip install requests
'''

import paramiko
import sys, os, json
from string import Template
from datetime import datetime
from paramiko import SSHClient, AutoAddPolicy, SFTPClient
import stat
import msal
import requests

# global set static variables
app_path = os.path.dirname(os.path.abspath(sys.argv[0]))
downloads_dir = os.path.join(app_path, 'downloads')
template_dir = os.path.join(app_path, 'templates')
config_file_path = os.path.join(app_path, 'config.json')

# global get config
with open(config_file_path, 'r') as f:
    config_json = json.load(f)


def getConfigByQuery(search_key, search_value, keys_to_include) -> dict:
    _found_config = False
    _configs = config_json['configs']
    for _config_data in _configs:
        if _config_data[search_key] == search_value:
            _found_config = True
            config_dict = getDictFromJson(json_data=_config_data, keys_to_include=keys_to_include)

    if _found_config == True:
        return config_dict
    else:
        raise ValueError(f"no config with: {search_key} = {search_value} found")


def getDictFromJson(json_data, keys_to_include) -> dict:
    config_dict = {key: json_data[key] for key in keys_to_include if key in json_data}
    return config_dict


def _getMailConfigByTitle(config_title) -> dict:
    _found_config = False
    _configs = config_json['configs']
    for _config_data in _configs:
        if _config_data['title'] == config_title:
            _found_config = True
            config_dict = {
                'title': _config_data.get('title'),
                'tenant_id': _config_data.get('tenant_id'),
                'client_id': _config_data.get('client_id'),
                'client_secret': _config_data.get('client_secret'),
                'scope': [_config_data.get('scope')]
            }
    if _found_config == True:
        return config_dict
    else:
        raise ValueError(f"no config with title: {config_title} found")



def _get_access_token(tenant_id, client_id, client_secret, scope) -> str | None:

    try:
        app = msal.ConfidentialClientApplication(
            authority=f"https://login.microsoftonline.com/{tenant_id}", 
            client_id=client_id, 
            client_credential=client_secret
        )
        result = app.acquire_token_silent(scopes=scope, account=None)
        if not result:
            # If no token in cache, get a new one from Azure AD
            result = app.acquire_token_for_client(scopes=scope)

        if "access_token" in result:
            return result["access_token"]
        else:
            print("Error acquiring token:", result.get("error_description"))
            return None
    
    except Exception as e:
        print("Exception while acquiring token:", str(e))
        return None


def _send_mail_msgraph_(sender, email_payload, access_token) -> bool:
    
    url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
    headers = {
        'Authorization': 'Bearer ' + str(access_token),
        'Content-Type': 'application/json'
    }

    # send request
    response = requests.post(url, headers=headers, json=email_payload)

    # capture response and return to the caller
    if str(response.status_code) != '202':
        return False
    return True






def _send_mail(sender, recipient, subject, content, access_token, content_type="HTML", save_to_sent_items=True) -> bool:
    
    if sender is not None and recipient is not None and subject is not None and content is not None and access_token is not None:

        #url = 'https://graph.microsoft.com/v1.0/users/' + str(user_ms_id) + '/sendMail'
        url = f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"
        #print(url)
        payload = json.dumps({
            "message": {
                "subject": str(subject),
                "body": {
                    "contentType": str(content_type),
                    "content": str(content)
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": str(recipient)
                        }
                    }
                ]
            },
            "saveToSentItems": str(save_to_sent_items).lower()
        })

        headers = {
            'Authorization': 'Bearer ' + str(access_token),
            'Content-Type': 'application/json'
        }

        # send request
        response = requests.request("POST", url, headers=headers, data=payload)

        # capture response and return to the caller
        #print(f"Status Code: {response.status_code}, Response: {response.text}")
        if str(response.status_code) != '202':
            return False
        return True
    else:
        return False


if __name__ == '__main__':
    
    # get the ftp connection settings
    properties = ['title', 'hostname', 'username', 'password', 'remotepath', 'sshHostKeyFingerprint']
    try:
        config = getConfigByQuery(search_key="title", search_value="ftp.zurrose.ch", keys_to_include=properties)
    except:
        print(f"error loading the settings for the ftp connection")
        sys.exit(1)

    # get email settings
    properties = ['title', 'tenant_id', 'client_id', 'client_secret', 'scope']
    try:
        mail_config = getConfigByQuery(search_key="title", search_value="msgraph_api_peter-it_email", keys_to_include=properties)
    except:
        print(f"error loading the settings for the mail connection")
        sys.exit(1)

    # load the email template
    template_file_path = os.path.join(template_dir, 'download_report.html')
    
    with open(template_file_path, 'r') as template_file:
        email_template = Template(template_file.read())


    ssh_client = SSHClient()
    ssh_client.set_missing_host_key_policy(AutoAddPolicy())

    try:
        print(f"connecting {config['hostname']}...")
        ssh_client.connect(hostname=config['hostname'], port=22, username=config['username'], password=config['password'])
        print("successfully connected")

        with ssh_client.open_sftp() as sftp:
            files = sorted(sftp.listdir_attr(config['remotepath']), key=lambda f: f.st_mtime, reverse=True)
            message = ""
            rows_html = ""
            rows_data = []
            for file in files:
                # Filter out directories; only download files
                if stat.S_ISREG(file.st_mode):
                    remote_file_path = f"{config['remotepath']}/{file.filename}"
                    remote_stat = sftp.stat(remote_file_path)
                    remote_mtime = remote_stat.st_mtime
                    remote_atime = remote_stat.st_atime
                    local_file_path = os.path.join(downloads_dir, file.filename)
                    # delte the file if it already exists
                    if os.path.exists(local_file_path):
                        os.remove(local_file_path)
                    sftp.get(remote_file_path, local_file_path)
                    os.utime(local_file_path, (remote_atime, remote_mtime))  # Set local file's access and modification times
                    modified = datetime.fromtimestamp(file.st_mtime)
                    #message = message + f"downloaded {file.filename} | {modified.strftime('%Y-%m-%d %H:%M:%S')}  \r\n"
                    message = message + f" {modified.strftime('%Y-%m-%d %H:%M:%S')} | {file.filename}  \r\n"
                    # rows_html += f"<tr><td>{file.filename}</td><td>{modified.strftime('%Y-%m-%d %H:%M:%S')}</td></tr>\n"
                    # rows_html += f'<tr style="border-bottom: 1px solid #f3f4f6;"><td style="padding: 12px 16px; color: #6b7280;">{file.filename}</td><td style="padding: 12px 16px; color: #6b7280;">{modified.strftime('%Y-%m-%d %H:%M:%S')}</td></tr>\n'
                    row_data = {"file_name": file.filename, "modified": modified.strftime('%Y-%m-%d %H:%M:%S')}
                    rows_data.append(row_data)

    except paramiko.AuthenticationException:
        print("Authentication failed, please check your credentials.")
    except paramiko.SSHException as e:
        print(f"SSH error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        if ssh_client:
            ssh_client.close()
            print("\nSSH connection closed.")
    
    print("\ndownloaded files:\n")
    print(message)

    
    # send the report via msgraph api

    formatted_datetime = datetime.now().strftime("%Y-%m-%d %H:%M")
    subject = f"Medbase Mailbox Inventory sFTP Download {formatted_datetime}"


    scopes = [mail_config['scope']] # scope has to be an array
    token = _get_access_token(tenant_id=mail_config['tenant_id'], client_id=mail_config['client_id'], client_secret=mail_config['client_secret'], scope=scopes)
    if token is None:
        print("Error getting Access Token:\n", token)
        sys.exit(1)
    
    
    # build the message body
    rows_html = ""
    for row in rows_data:
        rows_html += f'<tr style="border-bottom: 1px solid #f3f4f6;"><td style="padding: 12px 16px; color: #6b7280;">{row["file_name"]}</td><td style="padding: 12px 16px; color: #6b7280;">{row["modified"]}</td></tr>\n'

    data = {
        'title': subject,
        'table_rows': rows_html
    }

    rendered_html = email_template.substitute(data)


    email_payload = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",  # Tells Graph API to parse HTML instead of plain text
                "content": rendered_html
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": "joerg.peter@mexnet.ch"
                    }
                }
            ]
        },
        "saveToSentItems": "true"
    }

    success = _send_mail_msgraph_(sender="joerg.peter@peter-it.ch", email_payload=email_payload, access_token=token)
    if (success == True):
        print(f"\nEmail sent successfully")
    else:
        print(f"\nERROR sending Email")
