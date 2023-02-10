import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloEnterprisePropertyPane } from './HelloEnterprisePropertyPane';
import { AadHttpClient } from '@microsoft/sp-http';

export interface IHelloEnterpriseAdaptiveCardExtensionProps {
  title: string;
}

export interface IHelloEnterpriseAdaptiveCardExtensionState {
  daily: Number,
  weekly: Number
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloEnterprise_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloEnterprise_QUICK_VIEW';

export default class HelloEnterpriseAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloEnterpriseAdaptiveCardExtensionProps,
  IHelloEnterpriseAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloEnterprisePropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      daily: 0,
      weekly: 0,
     };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    // Hardcoded to ensure the API will return a value
    // let userEmail = this.context.pageContext.user.email
    var userEmail = "alex@groverale.onmicrosoft.com"

    return this._fetchDataFromSQLAPI(userEmail);
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloEnterprise-property-pane'*/
      './HelloEnterprisePropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HelloEnterprisePropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  private _fetchDataFromSQLAPI(userEmail: string): Promise<void> {
    return this.context.aadHttpClientFactory
      .getClient('91c459da-e9aa-41ae-b070-3d70592de2a2')
      .then(client => client.get(`https://ag-viva-connections-sql.azurewebsites.net/api/getcommission?userEmail=${userEmail}`, AadHttpClient.configurations.v1))
      .then(response => response.json())
      .then(commission => {
        this.setState({
          daily: commission.daily,
          weekly: commission.weekly
        });
      });
  }
}
