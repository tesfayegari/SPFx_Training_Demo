import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export default class SPService {
  
  constructor(private context: WebPartContext | ApplicationCustomizerContext) {
  }

  public getBirthdayItems() {
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    return this.context.spHttpClient.get(`${siteUrl}/_api/web/Lists/getbytitle('BDays')/items?$top=300&$select=*,Employee/Title,Employee/Name, Employee/EMail, Employee/MobilePhone, Employee/SipAddress, Employee/Department, Employee/JobTitle, Employee/FirstName, Employee/LastName, Employee/WorkPhone, Employee/UserName, Employee/Office, Employee/ID, Employee/Modified, Employee/Created&$expand=Employee`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<{ value: any[] }> => {
        return response.json();
      })
  }

  public getListItems(listName: string) {
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    return this.context.spHttpClient.get(`${siteUrl}/_api/web/Lists/getbytitle('${listName}')/items`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse): Promise<{ value: any[] }> => {
        return response.json();
      })
  }

  public getItemById(listName: string, itemId: number | string): Promise<any> {
    return new Promise<any>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      this.context.spHttpClient.get(`${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<any> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: any): void => {
          if (response.value.length === 0) {
            resolve(-1);
          }
          else {
            resolve(response);
          }
        });
    });
  }

  //only value of Title is going to be filled when Item is created
  public  createItem(listName: string): void {
    console.log({
      status: 'Creating item...',
      items: []
    });

    this.getListItemEntityTypeName(listName)
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        const body: string = JSON.stringify({
          '__metadata': {
            'type': listItemEntityTypeName
          },
          'Title': `Item ${new Date()}`
        });
        const webUrl = this.context.pageContext.web.absoluteUrl;
        return this.context.spHttpClient.post(`${webUrl}/_api/web/lists/getbytitle('${listName}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((item: any): void => {
        console.log({
          status: `Item '${item.Title}' (ID: ${item.Id}) successfully created`,
          items: []
        });
      }, (error: any): void => {
        console.log({
          status: 'Error while creating the item: ' + error,
          items: []
        });
      });
  }


  private getListItemEntityTypeName(listName: string): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {

      const webUrl = this.context.pageContext.web.absoluteUrl;
      this.context.spHttpClient.get(`${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          resolve(response.ListItemEntityTypeFullName);
        });
    });
  }

}