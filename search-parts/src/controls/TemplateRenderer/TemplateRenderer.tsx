import { isEqual } from "@microsoft/sp-lodash-subset";
import { LayoutRenderType } from "@pnp/modern-search-extensibility";
import * as DOMPurify from "dompurify";
import * as React from "react";
import { Constants } from "../../common/Constants";
import { DomPurifyHelper } from "../../helpers/DomPurifyHelper";
import { ISearchResultsTemplateContext } from "../../models/common/ITemplateContext";
import ITemplateRendererProps from "./ITemplateRendererProps";
import ITemplateRendererState from "./ITemplateRendererState";
import "./TemplateRenderer.scss";

// Need a root class to do not conflict with PnP Modern Search Styles.
const rootCssClassName = "iexp-modern-search";

export class TemplateRenderer extends React.Component<
   ITemplateRendererProps,
   ITemplateRendererState
> {
   private _domPurify: any;
   private _divTemplateRenderer: React.RefObject<HTMLDivElement>;

   constructor(props: ITemplateRendererProps) {
      super(props);

      this.state = {};

      this._domPurify = DOMPurify.default;

      this._domPurify.setConfig({
         ADD_TAGS: ["style", "#comment"],
         ADD_ATTR: ["target", "loading"],
         ALLOW_DATA_ATTR: true,
         ALLOWED_URI_REGEXP: Constants.ALLOWED_URI_REGEXP,
         WHOLE_DOCUMENT: true,
      });

      this._domPurify.addHook(
         "uponSanitizeElement",
         DomPurifyHelper.allowCustomComponentsHook
      );
      this._domPurify.addHook(
         "uponSanitizeAttribute",
         DomPurifyHelper.allowCustomAttributesHook
      );

      // Create an instance of the div ref container
      this._divTemplateRenderer = React.createRef<HTMLDivElement>();
   }

   public render() {
      return (
         <div
            className={rootCssClassName}
            ref={this._divTemplateRenderer}
         />
      );
   }

   public async componentDidMount() {
      await this.updateTemplate(this.props);

      if (this.props.analyticsService) {
         this.props.analyticsService.addResultHooks(
            this._divTemplateRenderer.current
         );
      }
   }

   public async componentDidUpdate(prevProps: ITemplateRendererProps) {
      if (
         !isEqual(prevProps.templateContent, this.props.templateContent) ||
         !isEqual(
            (prevProps.templateContext as ISearchResultsTemplateContext)
               .inputQueryText,
            (this.props.templateContext as ISearchResultsTemplateContext)
               .inputQueryText
         ) ||
         !isEqual(
            (prevProps.templateContext as ISearchResultsTemplateContext).data,
            (this.props.templateContext as ISearchResultsTemplateContext).data
         ) ||
         !isEqual(
            prevProps.templateContext.filters,
            this.props.templateContext.filters
         ) ||
         !isEqual(
            prevProps.templateContext.properties,
            this.props.templateContext.properties
         ) ||
         !isEqual(
            prevProps.templateContext.theme,
            this.props.templateContext.theme
         ) ||
         !isEqual(
            (prevProps.templateContext as ISearchResultsTemplateContext)
               .selectedKeys,
            (this.props.templateContext as ISearchResultsTemplateContext)
               .selectedKeys
         )
      ) {
         await this.updateTemplate(this.props);

         // Likely this is a change of page!
         if (this.props.analyticsService) {
            this.props.analyticsService.addResultHooks(
               this._divTemplateRenderer.current
            );
         }
      }
   }

   private async updateTemplate(props: ITemplateRendererProps): Promise<void> {
      let templateContent = props.templateContent;

      // Process the Handlebars template
      let template = await this.props.templateService.processTemplate(
         props.templateContext,
         templateContent,
         props.renderType
      );

      if (
         props.renderType == LayoutRenderType.Handlebars &&
         typeof template === "string"
      ) {
         // Sanitize the template HTML
         template = template
            ? this._domPurify.sanitize(`${template}`)
            : template;
         const templateAsHtml = new DOMParser().parseFromString(
            template as string,
            "text/html"
         );

         if (props.templateContext.properties.useMicrosoftGraphToolkit) {
            this.props.templateService.replaceDisambiguatedMgtElementNames(
               templateAsHtml
            );
         }

         // Get <style> tags from Handlebars template content and prefix all CSS rules by the Web Part instance ID to isolate styles
         const styleElements = templateAsHtml.getElementsByTagName("style");
         // let styles: string[] = [];
         // debugger;
         const allStyles = [];

         if (styleElements.length > 0) {
            // The prefix for all CSS selectors
            const elementPrefixId = `${this.props.templateService.TEMPLATE_ID_PREFIX}${this.props.instanceId}`;

            for (let i = 0; i < styleElements.length; i++) {
               const style = styleElements.item(i);

               let cssscope = style.dataset.cssscope as string;

               if (cssscope !== undefined && cssscope === "layer") {
                  allStyles.push(`@layer { ${style.innerText} }`);
               } else {
                  allStyles.push(
                     this.props.templateService.legacyStyleParser(
                        style,
                        elementPrefixId
                     )
                  );
               }
            }
         }

         if (
            this.props.templateContext.properties.useMicrosoftGraphToolkit &&
            this.props.templateService.MgtCustomElementHelper.isDisambiguated
         ) {
            allStyles.forEach((style, index) => {
               allStyles[index] = style.replace(
                  /mgt-/g,
                  `${this.props.templateService.MgtCustomElementHelper.prefix}-`
               );
            });
         }

         this._divTemplateRenderer.current.innerHTML = `<style id='st_template_renderer_${
            this.props.instanceId
         }'>
         ${allStyles.join(" ")}</style>
         <div id="${this.props.templateService.TEMPLATE_ID_PREFIX}${
            this.props.instanceId
         }">${templateAsHtml.body.innerHTML}
         </div>`;
      } else if (
         props.renderType == LayoutRenderType.AdaptiveCards &&
         template instanceof HTMLElement
      ) {
         this._divTemplateRenderer.current.innerHTML = "";
         this._divTemplateRenderer.current.appendChild(template as HTMLElement);
      }
   }
}
