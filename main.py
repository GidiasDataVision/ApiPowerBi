import os 
import msal
import requests
import json
import pandas as pd
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from dotenv import load_dotenv
load_dotenv()


def acquire_bearer_token(username, password, azure_tenant_id, client_id, scopes):
    try:
        app = msal.PublicClientApplication(client_id, authority=azure_tenant_id)
        result = app.acquire_token_by_username_password(username, password, scopes)
        return result["access_token"]
    except Exception as e:
        print(f'Falha ao autentificar: {e}')
        return None


def getData( url, token):

    url = url
    headers = {
        "Authorization": f"Bearer {token}"
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
       data = response.json()   
       return data['value']
    else:
        print(f'Status {response.status_code}')

def json_file(filename, lista):
    with open(filename, 'w', encoding='utf-8') as json_file:
        json.dump(lista, json_file, ensure_ascii=False, indent=4)

def acessarList(siteUrl,nameList):

    userName = os.getenv('username')
    userPassword = os.getenv('senha')

    ctx = ClientContext(siteUrl).with_credentials(UserCredential(userName, userPassword))
    list = ctx.web.lists.get_by_title(nameList)
    items = list.items.top(10000).get().execute_query()
    return items


def addItem(siteUrl,nameList,item_creation_info):
    
    userName = os.getenv('username')
    userPassword = os.getenv('senha')
    ctx = ClientContext(siteUrl).with_credentials(UserCredential(userName, userPassword))
    list = ctx.web.lists.get_by_title(nameList)
    list_item = list.add_item(item_creation_info)
    ctx.execute_query()

def main():

    #gera o Bearer tokem para conseguir acessar a api do power bi
    bearer_token = acquire_bearer_token(
        username= os.getenv('username'),
        password= os.getenv('senha'),
        azure_tenant_id="https://login.microsoftonline.com/"+ os.getenv('tenantId'),
        client_id= os.getenv('clientId'),
        scopes=["https://analysis.windows.net/powerbi/api/Workspace.Read.All"])
    print('gerado o Bearer token')
    workspaces = getData("https://api.powerbi.com/v1.0/myorg/groups",bearer_token)
    titleWorkspace = []
    allReports = [] 
    idReports = []
    workspacesnot = ['Admin monitoring', 'Budget Nortis', 'Budget Vibra', '57a47a90-483c-464b-93f6-174facec99f0']
    #acessado cada workspace e pegando o ID para fazer uma nova consulta para pegar os reports
    for workspace in workspaces:

        id = workspace['id']
        #add os titulos dos workspaces na lista vazia
        titleWorkspace.append(workspace['name'])
        #conectando na api do reports
        reports = getData(f'https://api.powerbi.com/v1.0/myorg/groups/{id}/reports',bearer_token)
        for report in reports:
            report['description'] = report.get('description', 'Sem Informação')

            #add os dados dos reports na lista vazia
        allReports.extend(reports)

    print('workspaces armazendos')
    print('reports armazendos')
    

    idReports.extend([item['id'] for item in allReports])


    #conectando na api LIST WORKSPACE
    itemsWorkspace =  acessarList('https://nortisinc1.sharepoint.com/sites/SistemadeControledeLicenas', 'DB_BI_WORKSPACE')
    print('conectado no lista workspace')
    titleItemsWorkspace = []
    for item in itemsWorkspace:
        titleItemsWorkspace.append(item.properties['field_4'])

    titleAddListWorkspace = [item for item in titleWorkspace if item not in titleItemsWorkspace]
    dataAddWorkspace = [workspace for workspace in workspaces if workspace['name'] in titleAddListWorkspace]

    print('armazenado os itens presente na Lista')
    print(f'itens a ADD na Lista:{dataAddWorkspace}')
    
    for itemAdd in dataAddWorkspace:
        item_creation_info = {
            'Title': itemAdd['type'], 
            'field_0': itemAdd['id'],
            'field_1': int(itemAdd['isReadOnly']),
            'field_2':int( itemAdd['isOnDedicatedCapacity']),
            'field_4': itemAdd['name'],
        }
        addItem('https://nortisinc1.sharepoint.com/sites/SistemadeControledeLicenas', 'DB_BI_WORKSPACE', item_creation_info)
    print('itens add')
    
    itensBodyMenu = acessarList('https://nortisinc1.sharepoint.com/sites/SistemadeControledeLicenas', 'DB_BI_BODY_MENU')
    print('conectado no lista Body Menu')
    dataMenuBody = []
    titleMenuBody = []
    ordem = 0

    for item in itensBodyMenu:
        
        title =  item.properties['name']
        if item.properties['ordem'] == None:
            ordem += 1
            data = {"name": item.properties['name'],"ordem":ordem}
        else:
            data = {"name": item.properties['name'],"ordem":item.properties['ordem']}
            
        titleMenuBody.append(title)
        dataMenuBody.append(data)

    titleItemsWorkspaceUpper = [item.upper() for item in titleWorkspace]
    titleAddBodyMenu = [item for item in titleItemsWorkspaceUpper if item.upper() not in titleMenuBody]
    dataAddallBodyMenu  = [item for item in  workspaces if item['name'].upper() in titleAddBodyMenu]
    maior_ordem = max(item['ordem'] for item in dataMenuBody if item['name'] != 'SOLICITAR PROJETO DE CONTRUÇÃO DE RELATORIO')
    
    print('armazenado os itens presente na Lista')
    print(f'itens a ADD na Lista:{dataAddallBodyMenu}')
    
    for itemAdd in dataAddallBodyMenu:
        maior_ordem += 1
        name = itemAdd['name'].upper()
        item_creation_info = {
            'name' :name,
            'ordem' : maior_ordem,
            'workspaceID':itemAdd['id']
        }    
        addItem('https://nortisinc1.sharepoint.com/sites/SistemadeControledeLicenas', 'DB_BI_BODY_MENU', item_creation_info)

    print('itens add')

    itemsReports =  acessarList('https://nortisinc1.sharepoint.com/sites/SistemadeControledeLicenas', 'BD_BI_ALL_REPORTS')
    print('conectado no lista Reports')
    idItemsReports = []
    for item in itemsReports:
        idItemsReports.append(item.properties['field_0'])
                     
    idAddListReports= [item for item in idReports if item not in idItemsReports]
    dataAddallReports = [item for item in allReports if item['id']  in idAddListReports]
    print('armazenado os itens presente na Lista')
    print(f'itens a ADD na Lista:{dataAddallReports}')
    print(datetime.today().strftime("%d/%m/%Y"))

    for itemAdd in dataAddallReports:
        item_creation_info = {
            'Title':itemAdd['name'],
            'field_0': itemAdd['id'], 
            'field_1':itemAdd['reportType'],
            'field_3': itemAdd['webUrl'],
            'field_4': itemAdd['embedUrl'],
            'field_5': int(itemAdd['isFromPbix']),
            'field_6': int(itemAdd['isOwnedByMe']),
            'field_7': itemAdd['datasetId'],
            'field_8': itemAdd['datasetWorkspaceId'],
            'field_9': json.dumps(itemAdd['users']), 
            'field_10': itemAdd['description'],

        }
        addItem('https://nortisinc1.sharepoint.com/sites/SistemadeControledeLicenas', 'BD_BI_ALL_REPORTS', item_creation_info)
    print('itens add')

   

if __name__ == "__main__":
    main()