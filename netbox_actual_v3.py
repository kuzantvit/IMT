# import sys
# sys.path.append('E:\\Positive Technologies\\MaxPatrol SIEM Agent\\modules\\PyEventCollector\\custompkg')
import re
import requests
import json
import codecs
from html.parser import HTMLParser

from requests.packages.urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

#Path to temperary files
TEMP_PATH = 'C:\\temp\\'
TEMP_FILE_PREFIX = 'hosts_'

NB_ADDRESS = 'specified in task settings'
NB_TOKEN = 'specified in task settings'
DVC_URL = '/api/dcim/devices/'
VM_URL = '/api/virtualization/virtual-machines/'

# Маппинг тэгов Netbox и групп активов. 
TAGS_GROUPS = [
    {
        'client_tag':'clouddc',
        'group':'CloudDC', # Это название должно присутствовать в описаниях самой группы и ее вложенных групп
        'service':'',
        'domains':[ # Предполагается, что названия групп по доменам соответствуют полю domain
            {'domain_tag':'clouddc',
            'domain':'clouddc.ru'},
			{'domain_tag':'customersclouddcru',
            'domain':'customers.clouddc.ru'},
            {'domain_tag':'infra',
            'domain':'infra.clouddc.ru'},
			{'domain_tag':'snegirsoftcom',
            'domain':'snegirsoft.com'},
            {'domain_tag':'rbsvpro',
            'domain':'rbsv.pro'},
			{'domain_tag':'infracdc',
            'domain':'infra.dc1.cdc'},
			{'domain_tag':'moncdc',
            'domain':'mon.cdc'},
			{'domain_tag':'netcdc',
            'domain':'net.cdc'}
        ]
    },
    {
        'client_tag':'sibintek',
        'group':'СИБИНТЕК', # Это название должно присутствовать в описаниях самой группы и ее вложенных групп
        'service':[
           {'svc_tag':'iaas',
            'svc_group':'СИБИНТЕК / IaaS'},
            {'svc_tag':'saas',
            'svc_group':'СИБИНТЕК / SaaS'},
        ],
        'domains':[ # Предполагается, что названия групп по доменам соответствуют полю domain
            {'domain_tag':'sibintek',
            'domain':'sibintek.ru'},
	    {'domain_tag':'sibinteksoftru',
            'domain':'sibintek-soft.ru'}
        ]
    }
]

#MaxPatrol SIEM connection
MPSIEM_CONNECTION = { 
        "core_url": "https://cdc-siem-core01.infra.clouddc.ru",
        "core_user": "specified in task settings",
        "core_pass": "specified in task settings",
        "auth_type": 0
    }
MPSIEM_UNMANAGED_ASSETGROUP = 'Common Cloud'

class AccessDenied(Exception):
    pass

def authenticate(address, login, password, new_password=None, auth_type=0):
    session = requests.session()
    session.verify = False

    response = print_response(session.post(
        address + ':3334/ui/login',
        json=dict(
            authType=auth_type,
            username=login,
            password=password,
            newPassword=new_password
        )
    ), check_status=False)

    if response.status_code != 200:
        raise AccessDenied(response.text)

    if '"requiredPasswordChange":true' in response.text:
        raise AccessDenied(response.text)

    return session, available_applications(session, address)

def available_applications(session, address):
    applications = print_response(session.get(
        address + ':3334/ptms/api/sso/v1/applications'
    )).json()

    return [
        app['id']
        for app in applications
        if is_application_available(session, app)
    ]

def is_application_available(session, app):
    if app['id'] == 'idmgr':
        modules = print_response(session.get(
            app['url'] + '/ptms/api/sso/v1/account/modules'
        )).json()

        return bool(modules)

    if app['id'] == 'mpx':
        return external_auth(
            session,
            app['url'] + '/account/login?returnUrl=/#/authorization/landing'
        )

def external_auth(session, address):

    response = print_response(session.get(address))

    if 'access_denied' in response.url:
        return False

    while '<form' in response.text:
        form_action, form_data = parse_form(response.text)

        response = print_response(session.post(form_action, data=form_data))

    return True

def parse_form(data):
    parser = HTMLParser()
    return re.search('action=[\'"]([^\'"]*)[\'"]', data).groups()[0], {
        item.groups()[0]: parser.unescape(item.groups()[1])
        for item in re.finditer(
            'name=[\'"]([^\'"]*)[\'"] value=[\'"]([^\'"]*)[\'"]',
            data
        )
    }

def print_response(response, check_status=True):
    if check_status:
        assert response.status_code == 200
    return response

def login_api(settings):
    settings['session'] = authenticate(settings['core_url'], settings[
        'core_user'], settings['core_pass'], auth_type=settings['auth_type'])[0]
        
def getChildren(group, name):
    if 'children' in group:
        for childrenGroups in group['children']:
            if childrenGroups['treePath'] == name:
                return childrenGroups['id']
            result = getChildren(childrenGroups, name)
            if result:
                return result

def getAssetGroupId(settings, group_name):
    res = settings['session'].get(
        settings['core_url'] + "/api/assets_temporal_readmodel/v2/groups/hierarchy?type=static").json()
    groupId = getChildren(res[0], group_name) 
    if groupId != None:
        return groupId
    return None

def importAssetList(settings, file_name, group_id):
    print("Starting Import")
    url = r"/api/assets_processing/v2/csv/import_operation?scopeId=00000000-0000-0000-0000-000000000005"
    upload_f = settings['session'].post(
        settings['core_url'] + url, files = {'upfile': ('host.csv', open(file_name, 'rb'), 'application/vnd.ms-excel')})
    r = json.loads(upload_f.text)
    print(file_name, group_id, r)

    data = {
            'groupsId': [group_id]
        }
    url = r"/api/assets_processing/v2/csv/import_operation/%s/start" % (r['id'])
    if r['totalRowsCount'] > 0:
        import_csv = settings['session'].post(
            settings['core_url'] + url, json=data)
        print(import_csv.status_code, import_csv.text)
    print("Stopping Import")
    
def check_hostname(hostname):
    hn = hostname.split(" ")[0]
    hn = hn.replace("(","-")
    hn = hn.replace("_","-")
    hn = hn.replace(")","")
    hn = hn.replace("[","")
    hn = hn.replace("]","")
    hn = hn.replace("/","-")
    return hn
    
 
def exportDataFromNetbox(host, url, token, temp_file, check_type, inf):
    s = requests.Session()
    headers = {'Authorization': 'Token ' + token,
    	"Accept":"application/json"
    }    
    address = 'https://' + host + url
    nb_link = 'https://' + host + re.search('api(.*)\/',url,re.IGNORECASE).group(1) + '/'    
    page_next = True
    reg = "^[a-z0-9-]*(\.[a-z0-9-]*)*\.[a-z]{2,}"
    
    try:                 
        if check_type == 'ordinary':
            r = s.get(address, headers = headers, verify=False)    

            with codecs.open(temp_file,"w","utf-8") as f:
                f.write('typealias;fqdn;hostname;ip;mac;isvirtual;uf_service_model;uf_netbox_url\n')
              
                while page_next:
                    r = s.get(address, headers = headers, verify=False)
                    j = json.loads(r.text)
                    for ip in j['results']:
                        if ip['primary_ip4'] != None and ip['name'] != None and ip['status']['value'] == 'active' and 'vc-res-lab' not in ip['tags'] and 'rosneft' not in ip['tags']:
                            hn = ip['name']
                            hn = check_hostname(hn)
                            hostname = re.search("^([a-zA-Z0-9-]*).*",hn,re.IGNORECASE).group(1)                            
                            fqdn = hn if (re.match(reg,hn,re.IGNORECASE) and re.search(reg,hn,re.IGNORECASE).group(0) == hn) else ""
                            typealias = re.search(".*(Windows|Linux).*",ip['platform']['name'],re.IGNORECASE).group(1) if ip['platform'] != None and ip['platform']['name'] != None and re.match(".*(Windows|Linux).*",ip['platform']['name'],re.IGNORECASE) else ""
                            txt = typealias.lower() + ';' + fqdn.lower() + ';' + hn.lower() + ';' + ip['primary_ip4']['address'].split("/")[0] + ';;;' + inf + ';' + nb_link + str(ip['id']) + '\n'
                            f.write(txt)
                                        
                    if j['next'] != None:
                        nxt_url = re.search('https:\/\/[0-9a-zA-Z.]*(.*)',j['next'],re.IGNORECASE).group(1)
                        address = 'https://' + host + nxt_url
                    else:
                        page_next = False
                f.close()
                
        if check_type == 'domains':
            r = s.get(address, headers = headers, verify=False)    

            with codecs.open(temp_file,"w","utf-8") as f:
                f.write('typealias;fqdn;hostname;ip;mac;isvirtual;uf_service_model;uf_netbox_url\n')
              
                while page_next:
                    r = s.get(address, headers = headers, verify=False)
                    j = json.loads(r.text)
                    for ip in j['results']:
                        if ip['primary_ip4'] != None and ip['name'] != None and ip['status']['value'] == 'active' and 'vc-res-lab' not in ip['tags'] and 'rosneft' not in ip['tags']:
                            hn = ip['name']
                            hn = check_hostname(hn)
                            hostname = re.search("^([a-zA-Z0-9-]*).*",hn,re.IGNORECASE).group(1)
                            fqdn = hn if (re.match(reg,hn,re.IGNORECASE) and re.search(reg,hn,re.IGNORECASE).group(0) == hn) else hn + '.' + inf if not re.match('.*'+inf+'.*',hn,re.IGNORECASE) else ""
                            hn = hn.replace("]","")
                            typealias = re.search(".*(Windows|Linux).*",ip['platform']['name'],re.IGNORECASE).group(1) if ip['platform'] != None and ip['platform']['name'] != None and re.match(".*(Windows|Linux).*",ip['platform']['name'],re.IGNORECASE) else ""
                            txt = typealias.lower() + ';' + fqdn.lower() + ';' + hn.lower() + ';' + ip['primary_ip4']['address'].split("/")[0] + ';;;;' + nb_link + str(ip['id']) + '\n'
                            f.write(txt)
                                        
                    if j['next'] != None:
                        nxt_url = re.search('https:\/\/[0-9a-zA-Z.]*(.*)',j['next'],re.IGNORECASE).group(1)
                        address = 'https://' + host + nxt_url
                    else:
                        page_next = False
                f.close()
                
        if check_type == 'unmanaged':
            r = s.get(address, headers = headers, verify=False)
            
            with codecs.open(temp_file,"w","utf-8") as f:
                f.write('typealias;fqdn;hostname;ip;mac;isvirtual;uf_service_model;uf_netbox_url\n')
              
                while page_next:
                    r = s.get(address, headers = headers, verify=False)
                    j = json.loads(r.text)
                    for ip in j['results']:
                        if ip['primary_ip4'] != None and ip['name'] != None and ip['status']['value'] == 'active' and 'vc-res-lab' not in ip['tags'] and 'rosneft' not in ip['tags']:
                            untagged = True
                            for c in inf:
                                if c['client_tag'] in ip['tags']:
                                    untagged = False 
                                for d in c['domains']:
                                    if d['domain_tag'] in ip['tags']:
                                        untagged = False
                            if untagged:  
                                hn = ip['name']
                                hn = check_hostname(hn)
                                hostname = re.search("^([a-zA-Z0-9-]*).*",hn,re.IGNORECASE).group(1)                                
                                fqdn = hn if (re.match(reg,hn,re.IGNORECASE) and re.search(reg,hn,re.IGNORECASE).group(0) == hn) else ""
                                typealias = re.search(".*(Windows|Linux).*",ip['platform']['name'],re.IGNORECASE).group(1) if ip['platform'] != None and ip['platform']['name'] != None and re.match(".*(Windows|Linux).*",ip['platform']['name'],re.IGNORECASE) else ""
                                txt = typealias.lower() + ';' + fqdn.lower() + ';' + hn.lower() + ';' + ip['primary_ip4']['address'].split("/")[0] + ';;;;' + nb_link + str(ip['id']) + '\n'
                                var_1_f = re.search(r'(sibintek.ru-)[^;]*', txt)
                                bad_dom = var_1_f.group(0)
                                txt=txt.replace(bad_dom, 'sibintek.ru')
                                f.write(txt)                
                    if j['next'] != None:
                        nxt_url = re.search('https:\/\/[0-9a-zA-Z.]*(.*)',j['next'],re.IGNORECASE).group(1)
                        address = 'https://' + host + nxt_url
                    else:
                        page_next = False
                f.close()              
                
    except Exception as e:
        return str(e)
		
def run(target, settings):
    savepoint = None
    *event_queue, has_more, savepoint = collect(target, settings, savepoint)

def collect(target, settings, savepoint):
    MPSIEM_CONNECTION['core_user'] = settings['second_credential']['login']
    MPSIEM_CONNECTION['core_pass'] = settings['second_credential']['password']
    login_api(MPSIEM_CONNECTION)
    
    
    SERVER_ADDRESS = target
    NB_TOKEN = settings['first_credential']['password']

    # Импорт по клиентам
    for c in TAGS_GROUPS:
        # Импорт всех клиентских активов в верхнеуровневые группы клиентов
        url_dvc = DVC_URL + '?tag=' + c['client_tag'].lower()
        url_vm = VM_URL + '?tag=' + c['client_tag'].lower()
        file_dvc = (TEMP_PATH + TEMP_FILE_PREFIX + c['group'] + '_dvc.csv').replace(" / ","")
        file_vm = (TEMP_PATH + TEMP_FILE_PREFIX + c['group'] + '_vm.csv').replace(" / ","")
        exportDataFromNetbox(SERVER_ADDRESS, url_dvc, NB_TOKEN, file_dvc, 'ordinary', '')  
        exportDataFromNetbox(SERVER_ADDRESS, url_vm, NB_TOKEN, file_vm, 'ordinary', '')        
        group_id = getAssetGroupId(MPSIEM_CONNECTION, c['group'])
        importAssetList(MPSIEM_CONNECTION, file_dvc, group_id)
        importAssetList(MPSIEM_CONNECTION, file_vm, group_id)
    
        # Импорт активов по IaaS, SaaS, и т.д.
        if c['service'] != '':
            for s in c['service']:
                url_dvc = DVC_URL + '?tag=' + c['client_tag'].lower() + '&tag=' + s['svc_tag'].lower()
                url_vm = VM_URL + '?tag=' + c['client_tag'].lower() + '&tag=' + s['svc_tag'].lower()
                file_dvc = (TEMP_PATH + TEMP_FILE_PREFIX + c['group'] + '_' + s['svc_group'] + '_dvc.csv').replace(" / ","")
                file_vm = (TEMP_PATH + TEMP_FILE_PREFIX + c['group'] + '_' + s['svc_group'] + '_vm.csv').replace(" / ","")
                exportDataFromNetbox(SERVER_ADDRESS, url_dvc, NB_TOKEN, file_dvc, 'ordinary', s['svc_tag'].lower())  
                exportDataFromNetbox(SERVER_ADDRESS, url_vm, NB_TOKEN, file_vm, 'ordinary', s['svc_tag'].lower())                
                group_id = getAssetGroupId(MPSIEM_CONNECTION, s['svc_group'])
                importAssetList(MPSIEM_CONNECTION, file_dvc, group_id)
                importAssetList(MPSIEM_CONNECTION, file_vm, group_id)
        
        # Импорт активов по доменам
        if c['domains'] != '':
            for d in c['domains']:
                url_dvc = DVC_URL + '?tag=' + d['domain_tag'].lower()
                url_vm = VM_URL + '?tag=' + d['domain_tag'].lower()
                file_dvc = (TEMP_PATH + TEMP_FILE_PREFIX + c['group'] + '_' + d['domain_tag'] + '_dvc.csv').replace(" / ","")
                file_vm = (TEMP_PATH + TEMP_FILE_PREFIX + c['group'] + '_' + d['domain_tag'] + '_vm.csv').replace(" / ","")
                exportDataFromNetbox(SERVER_ADDRESS, url_dvc, NB_TOKEN, file_dvc, 'domains', d['domain'])  
                exportDataFromNetbox(SERVER_ADDRESS, url_vm, NB_TOKEN, file_vm, 'domains', d['domain'])         
                group_id = getAssetGroupId(MPSIEM_CONNECTION, d['domain'])
                importAssetList(MPSIEM_CONNECTION, file_dvc, group_id)
                importAssetList(MPSIEM_CONNECTION, file_vm, group_id)
            
    # Импорт нераспределенных активов
    file_dvc = (TEMP_PATH + TEMP_FILE_PREFIX + 'Unmanaged_dvc.csv').replace(" / ","")
    file_vm = (TEMP_PATH + TEMP_FILE_PREFIX + 'Unmanaged_vm.csv').replace(" / ","")
    exportDataFromNetbox(SERVER_ADDRESS, DVC_URL, NB_TOKEN, file_dvc, 'unmanaged', TAGS_GROUPS) 
    exportDataFromNetbox(SERVER_ADDRESS, VM_URL, NB_TOKEN, file_vm, 'unmanaged', TAGS_GROUPS)     
    group_id = getAssetGroupId(MPSIEM_CONNECTION, MPSIEM_UNMANAGED_ASSETGROUP)
    importAssetList(MPSIEM_CONNECTION, file_dvc, group_id)
    importAssetList(MPSIEM_CONNECTION, file_vm, group_id)


    return False, savepoint

def main():
    
    run(TARGET, SETTINGS)	
    

if __name__ == "__main__":
    main()
