import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";

import styles from './LatestOrders.module.scss';
import { ILatestOrdersWebPartProps } from './ILatestOrdersWebPartProps';
import { IOrder, Region } from './IOrder';

export default class LatestOrdersWebPart extends BaseClientSideWebPart<ILatestOrdersWebPartProps> {
  private remotePartyLoaded: boolean = false;
  private orders: IOrder[];

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.latestOrders}">
      <iframe src="https://azure-ad-demo.neptune-preprod.bris.ac.uk/secure/user"
          style="display:none;"></iframe>
      <div class="ms-font-xxl">Username</div>
      <div class="loading"></div>
    </div>`;

    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement.querySelector(".loading"), "orders");

    this.domElement.querySelector("iframe").addEventListener("load", (): void => {
      this.remotePartyLoaded = true;
    });

    this.executeOrDelayUntilRemotePartyLoaded((): void => {
      this.context.httpClient.get("https://azure-ad-demo.neptune-preprod.bris.ac.uk/secure/user",
        HttpClient.configurations.v1, {
          credentials: "include"
        })
        .then((response: HttpClientResponse): Promise<any> => {
          
            //console.log(response.text()) ;

            return response.json();
            
          
        })
        .then((user: any): void => {
          console.log(user) ;
          //var username = user.split(":'")[1].split("'}")[0] ;
          var username = user.userid ;
          
          //this.user = user;
            this.context.statusRenderer.clearLoadingIndicator(
            this.domElement.querySelector(".loading").innerHTML = username);
          //this.renderData();
        })
        .catch((error: any): void => {

          this.context.statusRenderer.clearLoadingIndicator(
            this.domElement.querySelector(".loading"));
            //console.log(user) ;
          this.context.statusRenderer.renderError(this.domElement, "Error loading orders: " + (error ? error.message : ""));
        });
    });
  }




    private renderData(): void {
        if (this.orders) {
            const className = styles.number; // GLOBAL!
            const table: Element = this.domElement.querySelector(".data");
            table.removeAttribute("style");  //could use standard HTML5 'hidden' attribute instead of whole Style
            table.querySelector("tbody").innerHTML =
                this.orders.map(order =>
                    `<tr>
                        <td class="${className}">${order.id}</td>
                        <td class="${className}">${new Date(order.orderDate).toLocaleDateString()}</td>
                        <td>${order.region.toString()}</td>
                        <td>${order.rep}</td>
                        <td>${order.item}</td>
                        <td class="${className}">${order.units}</td>
                        <td class="${className}">$${order.unitCost.toFixed(2)}</td>
                        <td class="${className}">$${order.total.toFixed(2)}</td>
                      </tr>`
                ).join('');
        }
    }

  private executeOrDelayUntilRemotePartyLoaded(func: Function): void {
    if (this.remotePartyLoaded) {
      func();
    } else {
      setTimeout((): void => { this.executeOrDelayUntilRemotePartyLoaded(func); }, 100);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: []
    };
  }
}
