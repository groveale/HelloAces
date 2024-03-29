import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'HelloEnterpriseAdaptiveCardExtensionStrings';
import { IHelloEnterpriseAdaptiveCardExtensionProps, IHelloEnterpriseAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../HelloEnterpriseAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IHelloEnterpriseAdaptiveCardExtensionProps, IHelloEnterpriseAdaptiveCardExtensionState> {
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
      primaryText: `£${this.state.daily} Today 💸`,
      title: this.properties.title
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://www.bing.com'
      }
    };
  }
}
