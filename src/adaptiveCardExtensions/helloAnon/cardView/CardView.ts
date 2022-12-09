import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'HelloAnonAdaptiveCardExtensionStrings';
import { IHelloAnonAdaptiveCardExtensionProps, IHelloAnonAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../HelloAnonAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IHelloAnonAdaptiveCardExtensionProps, IHelloAnonAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IBasicCardParameters {
    return {
      primaryText: `${this.state.items.length} items from API`,
      title: this.properties.title
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://senddatatoace.azurewebsites.net/api/dummydata'
      }
    };
  }
}
