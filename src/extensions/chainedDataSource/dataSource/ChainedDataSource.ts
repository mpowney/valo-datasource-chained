import { BaseDataSourceProvider, IDataSourceData } from '@valo/extensibility';
import { IPropertyPaneGroup } from '@microsoft/sp-webpart-base';
import { HttpClient, SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { get } from '@microsoft/sp-lodash-subset';
import * as Msal from 'msal';
import * as strings from 'ChainedDataSourceApplicationCustomizerStrings';

const AAD_CONNECT_STORAGE_ENTITY = "ValoAadClientId";
const AAD_LOGIN_URL = "https://login.microsoftonline.com";
const LOADFRAME_TIMEOUT = 6000;

interface Map<T> {
    [K: string]: T;
}

export class ChainedDataSource extends BaseDataSourceProvider<IDataSourceData> {

    private clientId: string = '';
    private msalInstance: Msal.UserAgentApplication | undefined = undefined;
    private msalConfig: Map<Msal.Configuration>;
    private msalLoginRequest: Map<Map<Msal.AuthenticationParameters>>;
    private msalAuthResponse: Map<Map<Msal.AuthResponse | undefined>>;
    private propertyFieldCollectionData;
    private customCollectionFieldType;
    private tenantId: string;
    private tenantUrl: string;

    public async getData(): Promise<IDataSourceData> {

        const apiUrl = `${this.ctx.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('${AAD_CONNECT_STORAGE_ENTITY}')`;
        const storageEntity = await this.ctx.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1).then(storageData => storageData.json());

        if (storageEntity && storageEntity.Value) {
            this.clientId = storageEntity.Value;
        } else {
            throw `Storage entity ${AAD_CONNECT_STORAGE_ENTITY} was not found`;
        }

        let response: any = { items: [] };

        if (!(this.properties.apiUrl && this.properties.apiUrl.length)) {
            return response;
        }
        
        await this.initAuth();

        if (!this.msalAuthResponse) {
            console.log(`No authentication performed`);
        }
        else {
            console.log(this.msalAuthResponse);
        }

        for (let x = 0; x < this.properties.apiUrl.length; x++) {
            const chainItem = this.properties.apiUrl[x];
            const bearer = this.msalAuthResponse[chainItem.clientId || this.clientId] 
                && this.msalAuthResponse[chainItem.clientId || this.clientId][this.determineScope(chainItem)] 
                && this.msalAuthResponse[chainItem.clientId || this.clientId][this.determineScope(chainItem)].accessToken;

            if (!chainItem.authenticated) {

                const data = await this.ctx.httpClient.get(chainItem.apiUrl, HttpClient.configurations.v1, {
                    headers: {
                        "content-type": "application/json",
                        "accept": "application/json"
                    }
                });

                if (data && data.ok) {
                    response.items.push(await data.json());
                }
                else {
                    response.items.push({});
                }

            }
            else {

                if (bearer) {

                    const data = await this.ctx.httpClient.get(this.parseValue(chainItem.apiUrl, response), HttpClient.configurations.v1, {
                        headers: {
                            "Authorization": `Bearer ${this.msalAuthResponse[chainItem.clientId || this.clientId][this.determineScope(chainItem)].accessToken}`,
                            "content-type": "application/json",
                            "accept": "application/json"
                        }
                    });
    
                    if (data && data.ok) {
                        response.items.push(await data.json());
                    }
                    else {
                        response.items.push({});
                    }
    
                }
                else {
                    response.items.push({});
                }
            
            }
      
        }
        return response;

    }

    public getConfigProperties(): IPropertyPaneGroup[] {

        this.propertyFieldCollectionData = PropertyFieldCollectionData;
        this.customCollectionFieldType = CustomCollectionFieldType;

        const parametersControl = this.propertyFieldCollectionData("apiUrl", {
            key: "apiUrl",
            label: strings.WebPartPropertiesAPIURLsLabel,
            panelHeader: strings.WebPartPropertiesAPIURLPanelHeaderLabel,
            manageBtnLabel: strings.WebPartPropertiesAPIURLsButtonLabel,
            value: this.properties.apiUrl,
            enableSorting: true,
            fields: [
                {
                    id: "authenticated",
                    title: strings.WebPartPropertiesAuthenticatedColumnLabel,
                    type: this.customCollectionFieldType.boolean,
                },
                {
                    id: "apiUrl",
                    title: strings.WebPartPropertiesAPIURLColumnLabel,
                    type: this.customCollectionFieldType.string,
                    required: true
                },
                {
                    id: "method",
                    title: strings.WebPartPropertiesMethodColumnLabel,
                    type: this.customCollectionFieldType.dropdown,
                    required: true,
                    options: [ { key: 'GET', text: 'GET' }, { key: 'POST', text: 'POST' }, { key: 'OPTIONS', text: 'OPTIONS' }, { key: 'PATCH', text: 'PATCH' }, { key: 'PUT', text: 'PUT' }, { key: 'DELETE', text: 'DELETE' } ]
                },
                {
                    id: "clientId",
                    title: strings.WebPartPropertiesClientIdColumnLabel,
                    type: this.customCollectionFieldType.string,
                },
                {
                    id: "resource",
                    title: strings.WebPartPropertiesResourceColumnLabel,
                    type: this.customCollectionFieldType.string,
                }
            ]
        });
        return [
            {
              groupName: "Chained",
              groupFields: [
                parametersControl
              ],
              isCollapsed: false
            }
          ];
      
    }

    public parseValue(value: string, object: any) {

        const conditionalTokens = value.match(/\{\{[^\{]*?\}\}/gi);
        if (conditionalTokens !== null && conditionalTokens.length > 0) {
            for (let i = 0; i < conditionalTokens.length; i++) {
                const token = conditionalTokens[0].substring(2, conditionalTokens[0].length - 3);
                let condition = get(object, token, '');
                if (i === 0 && !condition) {
                    condition = "";
                    break;
                }
                console.log(`Previous value: ${value}`);
                value = value.replace(`{{${token}}}`, condition);
                console.log(`New value: ${value}`);
            }
        }

        return value;
    }

    private ensureMsalConfig(clientId: string, tenantScopedAuth: boolean) {

        this.msalConfig = this.msalConfig || {};
        this.msalConfig[clientId] = this.msalConfig[clientId] || {
            auth: {
                clientId: this.clientId,
                authority: `${AAD_LOGIN_URL}/${tenantScopedAuth ? this.tenantId : "common"}`,
                redirectUri: `${this.tenantUrl}/_layouts/images/blank.gif`,
            },
            system: {
                loadFrameTimeout: LOADFRAME_TIMEOUT
            }
        };

    }

    private ensureMsalLoginRequest(clientId: string, scope: string, loginName: string) {

        this.msalLoginRequest = this.msalLoginRequest || {};
        this.msalLoginRequest[clientId] = this.msalLoginRequest[clientId] || {};
        this.msalLoginRequest[clientId][scope] = this.msalLoginRequest[clientId][scope] || {
            scopes: [scope],
            loginHint: this.ctx.pageContext.user.loginName,
        };

    }

    private ensureMsalAuthResponse(clientId: string, scope: string, authResponse: Msal.AuthResponse) {

        this.msalAuthResponse = this.msalAuthResponse || {};
        this.msalAuthResponse[clientId] = this.msalAuthResponse[clientId] || {};
        this.msalAuthResponse[clientId][scope] = this.msalAuthResponse[clientId][scope] || authResponse;

    }

    private determineScope(chainItem: any) {
        const defaultAadScope = (chainItem.apiUrl && chainItem.apiUrl.indexOf("https://") > -1) ? `${chainItem.apiUrl.substring(0, chainItem.apiUrl.indexOf("/", 8))}/user_impersonation` : '';
        return chainItem.resource || defaultAadScope;
    }
    
    public async initAuth() {

        this.tenantId = this.ctx.pageContext.aadInfo.tenantId;
        this.tenantUrl = this.ctx.pageContext.site.absoluteUrl.replace(this.ctx.pageContext.site.serverRelativeUrl, "");

        for (let x = 0; x < this.properties.apiUrl.length; x++ ) {

            const chainItem: any = this.properties.apiUrl[x];
            this.ensureMsalConfig(chainItem.clientId || this.clientId, true);
            this.ensureMsalLoginRequest(chainItem.clientId || this.clientId, this.determineScope(chainItem), this.ctx.pageContext.user.loginName);
            console.log(`DynamicsDataSource scope = ${JSON.stringify([chainItem.clientId || this.clientId, this.determineScope(chainItem), this.ctx.pageContext.user.loginName])}`);

        }


        for (const key in this.msalLoginRequest) {
            
            const loginRequests = this.msalLoginRequest[key];
            this.msalInstance = new Msal.UserAgentApplication(this.msalConfig[key]);
    
            if (this.msalInstance) {
    
                this.msalInstance.handleRedirectCallback((error: any, response: any) => {
                    // handle redirect response or error
                    if (error) {
                        console.log(`Error: ${error.errorMessage}`);
                    } else if (response) {
                        console.log(`Response from MSAL: ${response.account}`);
                    }
                });
            
            }
    
            console.log(`Calling acquireTokenSilent()`);
    
            try {
                for (const loginRequestKey in loginRequests) {

                    this.ensureMsalAuthResponse(key, loginRequestKey, await this.msalInstance.acquireTokenSilent(loginRequests[loginRequestKey]));
                }
                
            } catch (err) {
                console.log(err);
            }

        }


    }

}
