import {
   IReadonlyTheme,
   ThemeChangedEventArgs,
   ThemeProvider,
} from "@microsoft/sp-component-base";
import { Log } from "@microsoft/sp-core-library";
import { isEqual } from "@microsoft/sp-lodash-subset";
import {
   IPropertyPaneGroup,
   PropertyPaneLabel,
   PropertyPaneLink,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { PropertyPaneWebPartInformation } from "@pnp/spfx-property-controls/lib/PropertyPaneWebPartInformation";
import * as commonStrings from "CommonStrings";
import { IBaseWebPartProps } from "../models/common/IBaseWebPartProps";
import { ExtensibilityService } from "../services/extensibilityService/ExtensibilityService";
import IExtensibilityService from "../services/extensibilityService/IExtensibilityService";

/**
 * Generic abstract class for all Web Parts in the solution
 */
export abstract class BaseWebPart<
   T extends IBaseWebPartProps
> extends BaseClientSideWebPart<IBaseWebPartProps> {
   /**
    * Theme variables
    */
   protected _themeProvider: ThemeProvider;
   protected _themeVariant: IReadonlyTheme;

   /**
    * The Web Part properties
    */
   // @ts-ignore: redefinition
   protected properties: T;

   /**
    * The data source service instance
    */
   protected extensibilityService: IExtensibilityService = undefined;

   constructor() {
      super();
   }

   /**
    * Initializes services shared by all Web Parts
    */
   protected async initializeBaseWebPart(): Promise<void> {
      // Get service instances
      this.extensibilityService =
         this.context.serviceScope.consume<IExtensibilityService>(
            ExtensibilityService.ServiceKey
         );

      // Initializes them variant
      this.initThemeVariant();

      return;
   }

   /**
    * Returns common information groups for the property pane
    */
   protected getPropertyPaneWebPartInfoGroups(): IPropertyPaneGroup[] {
      return [
         {
            groupName: commonStrings.General.About,
            groupFields: [
               PropertyPaneWebPartInformation({
                  description: `<span>${commonStrings.General.Authors}: <ul style="list-style-type: none;padding: 0;"><li><img width="16px" src="data:image/png;base64,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"/><a href="https://informationexp.com/">Information EXP</a></li><li>Based upon PnP Modern Search Parts</li></ul></span>`,
                  key: "authors",
               }),
               PropertyPaneLabel("", {
                  text: `${commonStrings.General.Version}: ${
                     this && this.manifest.version ? this.manifest.version : ""
                  }`,
               }),
               PropertyPaneLabel("", {
                  text: `${commonStrings.General.InstanceId}: ${this.instanceId}`,
               }),
            ],
         },
         {
            groupName: commonStrings.General.Resources.GroupName,
            groupFields: [
               PropertyPaneLink("", {
                  target: "_blank",
                  href: this.properties.documentationLink,
                  text: commonStrings.General.Resources.Documentation,
               }),
            ],
         },
      ];
   }

   /*
    * Get the parent control zone for the current Web Part instance
    */
   protected getParentControlZone(): HTMLElement {
      // 1st attempt: use the DOM element with the Web Part instance id
      let parentControlZone = document.getElementById(this.instanceId);

      if (!parentControlZone) {
         // 2nd attempt: Try the data-automation-id attribute as we suppose MS tests won't change this name for a while for convenience.
         parentControlZone = this.domElement.closest(
            `div[data-automation-id='CanvasControl'], .CanvasControl`
         );

         if (!parentControlZone) {
            // 3rd attempt: try the Control zone with ID
            parentControlZone = this.domElement.closest(
               `div[data-sp-a11y-id="ControlZone_${this.instanceId}"]`
            );

            if (!parentControlZone) {
               Log.warn(
                  this.manifest.alias,
                  `Parent control zone DOM element was not found in the DOM.`,
                  this.context.serviceScope
               );
            }
         }
      }

      return parentControlZone;
   }

   /**
    * Initializes theme variant properties
    */
   private initThemeVariant(): void {
      // Consume the new ThemeProvider service
      this._themeProvider = this.context.serviceScope.consume(
         ThemeProvider.serviceKey
      );

      // If it exists, get the theme variant
      this._themeVariant = this._themeProvider.tryGetTheme();

      // Register a handler to be notified if the theme variant changes
      this._themeProvider.themeChangedEvent.add(
         this,
         this._handleThemeChangedEvent.bind(this)
      );
   }

   /**
    * Update the current theme variant reference and re-render.
    * @param args The new theme
    */
   private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
      if (!isEqual(this._themeVariant, args.theme)) {
         this._themeVariant = args.theme;
         this.render();
      }
   }
}
