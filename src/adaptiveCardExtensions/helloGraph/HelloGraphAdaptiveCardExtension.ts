import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloGraphPropertyPane } from './HelloGraphPropertyPane';

export interface IHelloGraphAdaptiveCardExtensionProps {
  title: string;
}

export interface IHelloGraphAdaptiveCardExtensionState {
  name?: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloGraph_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloGraph_QUICK_VIEW';

export default class HelloGraphAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloGraphAdaptiveCardExtensionProps,
  IHelloGraphAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloGraphPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return this._fetchData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloGraph-property-pane'*/
      './HelloGraphPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HelloGraphPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  private _fetchData(): Promise<void> {
    return this.context.msGraphClientFactory
      .getClient("3")
      .then(client => client.api('me').get())
      .then((user) => {
        this.setState({
          name: user.displayName
        });
      });
  }
}
