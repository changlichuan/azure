import getopt
import sys
import json
import requests
import csv

MSLOGIN_URL = "https://login.microsoftonline.com/"
MSLOGIN_URL_SUFFIX = "/oauth2/v2.0/token"
MSUSER_URL = "https://graph.microsoft.com/v1.0/Users/"
MSAPPLICATION_URL = "https://graph.microsoft.com/v1.0/Applications/"

STATUS_ASSIGNED='assigned';
STATUS_UNCHANGE='existing';
STATUS_FAILED='failed';
STATUS_NOTFOUND = 'user_not_found'

# Open CSV with list of email addresses (one per line)
def getUsers(csv_filename) :
    users = [];
    with open(csv_filename, 'rt') as f:
        reader = csv.reader(f)
        for row in reader:
            #print ('Email Address Read: '+ str(row))
            users.append(row[0])
    return users;

# Obtain a token valid for 1 hour
def getToken(tenant_id,client_id,client_secret):
    payload = {'client_id':client_id,'client_secret':client_secret,'grant_type':'client_credentials','scope':'https://graph.microsoft.com/.default'}
    url = MSLOGIN_URL+tenant_id+MSLOGIN_URL_SUFFIX;
    # print ('SCIM URI SENT: ' + url)
    tokenData={};
    tokenResponse = requests.post(url, data=payload)
    if str(tokenResponse)=='<Response [200]>':
        tokenData = tokenResponse.json()
        if 'access_token' in tokenData :
            token = tokenData['access_token']
            #print ('Token obtained')
            return token
    return '';

def getRoleID(app_id,access_token):
    url = MSAPPLICATION_URL;
    appData = {};
    appResponse = requests.get(url, headers=getHeaders(access_token)) 
    if str(appResponse)=='<Response [200]>':
        appData = appResponse.json()
        appList = [];
        if 'value' in appData :
            appList = appData['value'];
            for app in appList :
                #found workplace integration app
                if ('appId' in app) and (str(app['appId']) == app_id):
                    roles = app['appRoles'];
                    for role in roles :
                        if ('displayName' in role) and (str(role['displayName'])=='User'):
                            return str(role["id"]);
    return '';


def getHeaders(access_token,content_type=''):
  headers = {'Authorization': 'Bearer ' + access_token}
  if len(content_type) > 0 :
      headers['Content-Type'] = content_type;
  #print(headers)
  return headers

def getUserID(user_email,access_token):
    url = MSUSER_URL+user_email;
    userData = {};
    userIDResponse = requests.get(url, headers=getHeaders(access_token)) 
    if str(userIDResponse)=='<Response [200]>':
        userData = userIDResponse.json()
        if 'id' in userData :
            user_id = userData['id']
            return user_id
    return '';

# Assign a user to the Workplace from Facebook integration
def assignUser(user_email,app_obj_id,role_id,access_token):

    user_id = getUserID(user_email,access_token);
    url = 'https://graph.microsoft.com/v1.0/Users/'+user_email+'/appRoleAssignments'
    if len(user_id) > 0 :
        #print('user_id found:'+user_id);
        payload = {'principalId':user_id,'resourceId':app_obj_id,'appRoleId':role_id}
        
        assignResponse = requests.post(url, json=payload, headers=getHeaders(access_token,'application/json'))
        if str(assignResponse) == '<Response [201]>':
            return STATUS_ASSIGNED
        elif str(assignResponse) == '<Response [400]>' :
            return STATUS_UNCHANGE
        else : 
            #print('Error in assignment: '+str(assignResponse.content))
            return STATUS_FAILED
    else : return STATUS_NOTFOUND        

def printUsage():
  print('aad-assign.py -f <csv filename> -t <tenant_id> -a <app_id> -o <app_obj_id> -s <script_id> -c <client_secret>')



def main(argv):
    try:
        options, args = getopt.getopt(argv, 'hf:t:a:o:s:c:')
    except getopt.GetoptError:
        printUsage()
        sys.exit(-1)

    tenant_id = ''   #tenant id of your AzureAD or primary domain
    app_id = ''      #workplace integration application id
    app_obj_id = ''  #workplace integration object id
    script_id = ''   #Application ID of registered app (this script)
    client_secret = '' #client_secret of registered app
    filename = ''                               #filename of csv containing all users' email, 1 per line

    for option, arg in options:
        if option == '-h':
            printUsage()
            sys.exit()
        elif option == '-t':
            tenant_id = arg
        elif option == '-a':
            app_id = arg
        elif option == '-o':
            app_obj_id = arg
        elif option == '-s':
            script_id = arg
        elif option == '-c':
            client_secret = arg
        elif option == '-f':
            filename = arg

    if script_id == '' or client_secret == '' or filename == '':
        printUsage()
        sys.exit(-1)


    #Main processing
    success_list = [];
    failed_list = [];
    not_found_list = [];
    existing_count = 0;

    token = getToken(tenant_id,script_id,client_secret)
    role_id = '';

    if len(token)>0 :
        u_list = getUsers(filename);
        total = len(u_list);
        if total == 0 :
            print('Empty CSV file')
            return;

        role_id = getRoleID(app_id,token)
        if len(role_id)==0 :
            print('Unable to retrive user role, please ensure Application ID of Azure Workplace Integration is filled correctly into app_id')
            return;

        for user in u_list :
            result = assignUser(user,app_obj_id,role_id,token);
            if result == STATUS_ASSIGNED :
                success_list.append(user);
            elif result == STATUS_FAILED :
                failed_list.append(user);
            elif result == STATUS_UNCHANGE :
                existing_count  = existing_count+1;
            elif result == STATUS_NOTFOUND :
                not_found_list.append(user);

            percent = int((u_list.index(user)+1)*100 / total);
            print('    Progress: %d %%'%(percent), end='\r')
        print();


        # Print the final outcome
        print('TOTAL: %d users' % (total))
        print('%d users newly assigned' % (len(success_list)));
        for _u in success_list :
            print('   '+_u);
        if existing_count>0 :
            print('%d users already assigned in the past' % (existing_count));
        if len(failed_list)>0 :
            print('%d users encountered error in assignment' % (len(failed_list)));
            for _u in failed_list :
                print('   '+_u);
        if len(not_found_list)>0 :
            print('%d users does not exist in your AzureAD' % (len(not_found_list)));
            for _u in not_found_list :
                print('   '+_u);
    else :
        print('Error retrieving AzureAD token, please double check permission and client secret parameters.')


if __name__ == '__main__':
    main(sys.argv[1:])


