import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloAnonPropertyPane } from './HelloAnonPropertyPane';
import { HttpClient } from '@microsoft/sp-http';

export interface IHelloAnonAdaptiveCardExtensionProps {
  title: string;
}

export interface IHelloAnonAdaptiveCardExtensionState {
  items: object[]
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloAnon_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloAnon_QUICK_VIEW';

export default class HelloAnonAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloAnonAdaptiveCardExtensionProps,
  IHelloAnonAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloAnonPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { 
      items: []
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchData()
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloAnon-property-pane'*/
      './HelloAnonPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HelloAnonPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  // remeber to add CORS to the API 
  private _fetchData(): Promise<void> {
    return this.context.httpClient
    .get('https://senddatatoace.azurewebsites.net/api/dummydata', HttpClient.configurations.v1)
    .then((res: any): Promise<any> => {
      return res.json();
    })
    .then((response: any): void => {
      this.setState({
        items: response.items
      });
    });
  }

}
