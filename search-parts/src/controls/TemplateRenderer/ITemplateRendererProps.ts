import { LayoutRenderType } from "@pnp/modern-search-extensibility";
import { IAnalyticsService } from "../../services/analyticsService/IAnalyticsService";
import { ITemplateService } from "../../services/templateService/ITemplateService";

interface ITemplateRendererProps {
   instanceId: string;

   /**
    * The template context
    */
   templateContext: any;

   /**
    * The Handlebars raw template content for a single item
    */
   templateContent: string;

   /**
    * A template service instance
    */
   templateService: ITemplateService;

   /**
    * The layout render type (Handlebars, Adaptive Cards, etc.)
    */
   renderType: LayoutRenderType;

   /**
    * INFO EXP: The analytics service instance
    */
   analyticsService?: IAnalyticsService;
}

export default ITemplateRendererProps;
