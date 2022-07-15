from turtle import delay
import requests
from datetime import datetime
import buildcustom
import ast
import copy

# api key is provided by clickup
apikey = input("Enter api key should start with pk_ ")
#team is the first value in the url https://app.clickup.com/########/v
team = input("Enter Workspace ID#")
#view is the last part of the url when viewing a list or a table
view = input("Enter View ID:")

#The name of the xls report to be created. 
file = input("file name:")

if input("would you like Base columns (URL, Space, folder, list, id)?  Y for yes: ").upper() == 'Y':
    extra = True
else:
    extra = False

# if new custom fields are created, the custom field json should be rebuilt, this can take a while depending on how many custom fields and unique values exist
if input("would you like to refresh custom field metadata(slow) ?  Y for yes: ").upper() == 'Y':
    print('rebuilding custom fields meta data - this can take a minute')
    custom = buildcustom.customlist()
else:
    cf = open('custom_meta.json', "r")
    ct = cf.read().replace('\n', '')
    custom = ast.literal_eval(ct)
    cf.close()

badlists = [] # place items you would like to not report here - Perhaps look up lists etc...
spaces = {}

payload={}                   
headers = {'Authorization': apikey}
f = open(file+".csv", "w",encoding="utf-8")
print('creating csv file '+file)
url = f'https://api.clickup.com/api/v2/team/{team}/space?archived=false'

while True:
    response = requests.request("GET", url, headers=headers, data=payload)
    if response.status_code == 200:
        break
    elif response.status_code == 429:
        delay(2)
    else:
        print('job failed after receiving invalid response from API response ='+str(response.status_code))
        quit()
temp = response.json()

print('building array of spaces')
for space in temp['spaces']: 
    spaces[space['id']] = space['name']

url = "https://api.clickup.com/api/v2/view/"+view

while True:
    response = requests.request("GET", url, headers=headers, data=payload)
    if response.status_code == 200:
        break
    elif response.status_code == 429:
        delay(2)
    else:
        print('job failed after receiving invalid response from API response ='+str(response.status_code))
        quit()
temp = response.json()

if extra:
    header = 'URL, Space, folder, list, id, name, '
else:
    header = 'name,'

hdr = {}
print('building column headers ')
firsttime = True
i = -1
IncludeSubtasks = False
if temp['view']['settings']['show_subtasks'] > 1:
    IncludeSubtasks = True
    

for column in temp['view']['columns']['fields']: 
    if not column['hidden'] :
        i += 1
        if not firsttime:
          header += ',' 
        if column['field'][:2]=='cf': #custom fields need to pull the name from custom dict
            header += custom[column['field'][3:]]['name']
            hdr[i] = {"from": 2, "id": column['field'][3:], "type":custom[column['field'][3:]]['type'] }
        else:
            header += column['field']
            hdr[i] = {"from": 1, "id": column['field']}
        firsttime = False

url = "https://api.clickup.com/api/v2/view/"+view+"/task?page="

i = 0

f.write(header+'\n')    

while True:
    print('pulling page '+str(i)+' from view '+view )
    while True:
        response = requests.request("GET", url+str(i), headers=headers, data=payload)
        if response.status_code == 200:
            break
        elif response.status_code == 429:
            delay(2)
        else:
            print('job failed after receiving invalid response from API response ='+str(response.status_code))
            quit()
    temp = response.json()
    if "err" in temp:
        print("error!  " + temp["err"])
        quit()
    tasks = temp['tasks']
    if IncludeSubtasks:
        templist = []
        for task in tasks:
            ttask = copy.deepcopy(task)
            if ttask['subtasks_count'] > 0:
                del ttask['subtasks']
                templist.append(ttask)
                for sub in task['subtasks']:  
                    sub['name'] = 'Subtask-'+sub['name'].replace(",", "").replace('•','-')  
                    templist.append(sub)
            else:
                templist.append(ttask)
        tasks = copy.deepcopy(templist)
    for task in tasks:
        if task['list']['id'] not in badlists:
            if extra:
                f.write(task['url']+',')
                f.write(spaces[task['space']['id']]+',')
                if task['folder']['hidden']:
                    f.write(',')
                else:
                    f.write(task['folder']['name']+',')

                f.write(task['list']['name']+',')
                f.write(task['id']+',')
            f.write(task['name'].replace(",", "").replace('•','-')+',')     
            firsttime = True
            for t in hdr:
                if not firsttime:
                    f.write(',')
                else:
                    firsttime = False

                h = hdr[t]
                if h['from']==1:
                    if h['id'] == 'status':
                        f.write(task['status']['status'])
                    elif h['id'] == 'createdBy':
                        f.write(task['creator']['username'])  
                    elif h['id'] == 'assignee':
                        if len(str(task['assignees'])) > 0:
                            for assigned in task['assignees']:
                                if assigned['username'] is not None:
                                    f.write(assigned['username'])
                        
                    elif h['id'] == 'tags':
                        for tag in task['tags']:
                            f.write(tag['name'])
                        
                    elif h['id']=='priority':
                        if task['priority'] is None:
                            pass    
                        else:
                            f.write(task['priority']['priority'])
                    elif h['id'] == 'dueDate':
                        if task['due_date'] is None:
                            pass
                        else:
                            ts = int(task['due_date'])/1000
                            f.write(datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'))
                    elif h['id'] == 'startDate':
                        if task['start_date'] is None:
                            pass
                        else:
                            ts = int(task['start_date'])/1000
                            f.write(datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'))
                    elif h['id'] == 'dateUpdated':
                        if task['date_updated'] is None:
                            pass
                        else:
                            ts = int(task['date_updated'])/1000
                            f.write(datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'))
                    elif h['id'] == 'dateClosed':
                        if task['date_closed'] is None:
                            pass
                        else:
                            ts = int(task['date_closed'])/1000
                            f.write(datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'))
                    elif h['id'] == 'dateCreated':
                        if task['date_created'] is None:
                            pass
                        else:
                            ts = int(task['date_created'])/1000
                            f.write(datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S'))
                    elif h['id'] == 'timeEstimate':
                        if task['time_estimate'] is None:
                            pass
                        else:
                            min = int(task['time_estimate'])/60000
                            f.write(str(min))
                    elif h['id'] == 'timeLogged':
                        if task['time_logged'] is None:
                            pass
                        else:
                            min = int(task['time_logged'])/60000
                            f.write(str(min))
                    elif h['id'] == 'incompleteCommentCount':
                        url1 = "https://api.clickup.com/api/v2/task/"+task['id']+"/comment"
                        while True:
                            response = requests.request("GET", url1, headers=headers, data=payload)
                            if response.status_code == 200:
                                break
                            elif response.status_code == 429:
                                delay(2)
                            else:
                                print('job failed after receiving invalid response from API response ='+str(response.status_code))
                        temp1 = response.json()
                        ccount = 0
                        for c in temp1['comments']:
                            if 'resolved' in c:
                                if c['resolved'] is True:
                                    ccount += 1 
                        f.write(str(ccount))
                       
                    elif h['id'] == 'latestComment':
                        url1 = "https://api.clickup.com/api/v2/task/"+task['id']+"/comment"
                        while True:
                            response = requests.request("GET", url1, headers=headers, data=payload)
                            if response.status_code == 200:
                                break
                            elif response.status_code == 429:
                                delay(2)
                            else:
                                print('job failed after receiving invalid response from API response ='+str(response.status_code))
                        temp1 = response.json()
                        if len(temp1['comments']) > 0:
                            txt = temp1['comments'][0]['comment_text'].replace(",", "")
                            text1 = txt.replace("\n","")
                            f.write(text1)
                    elif h['id'] == 'lists':
                        if task['list'] is None:
                            pass
                        else:
                            f.write(task['list']['name'])
                            
                    else:
                        f.write(task[h['id']])
                else:
                    for custom in task['custom_fields']:
                        if custom['id']== h['id']:
                            if 'value' in custom:
                                if custom['type']=='date':
                                    ts = int(int(custom['value']))/1000
                                    f.write(datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')) 
                                elif custom['type'] == 'drop_down':
                                    if list(custom['type_config']['options'][custom['value']]['name']).count('|')==3:
                                        fields = custom['type_config']['options'][custom['value']]['name'].split('|')
                                        f.write(fields[0].replace(",", "")+",")
                                        f.write(fields[1].replace(",", "")+",")
                                        f.write(fields[2].replace(",", "")+",")
                                        f.write(fields[3].replace(",", ""))
                                    else:
                                        f.write(custom['type_config']['options'][custom['value']]['name'].replace(',',''))
                                else:
                                    tt = custom['value'].replace('\n','')
                                    ttt = tt.replace(',','')
                                    tttt = ttt.replace('\r','')
                                    f.write(tttt)
                            else:
                                if list(custom['name']).count('|')==3:
                                    f.write(',,,')       

            f.write('\n') 
                         
    if temp['last_page']:
        print('last page found - closing file')
        f.close()
        break

    i += 1

 



