import requests

def customlist(apikey, team):
    payload={}
    headers = {'Authorization': apikey}

    url =f"https://api.clickup.com/api/v2/{team}/36089457/space?archived=false"
    response = requests.request("GET", url, headers=headers, data=payload)
    temp = response.json()
    s = temp['spaces'] 
    i = 1
    lists = []

    for space in s:
        url ="https://api.clickup.com/api/v2/space/"+space['id']+"/list?archived=false"
        response = requests.request("GET", url, headers=headers, data=payload)
        temp = response.json()
        l = temp['lists'] 
        for id in l:
            lists.append(id['id'])

    custom = {}

    for id in lists:
        url = 'https://api.clickup.com/api/v2/list/'+id+'/field'
        response = requests.request("GET", url, headers=headers, data=payload)
        temp = response.json()
        c = temp['fields'] 
        for id in c:
            custom[id['id']] = {"name":id['name'],"type":id['type']}
    f = open('custom_meta.json', "w")
    f.write(str(custom))
    return custom        


