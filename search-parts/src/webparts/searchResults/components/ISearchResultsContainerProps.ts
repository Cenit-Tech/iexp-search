import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import {
   IDataContext,
   IDataFilterResult,
   IDataSource,
   LayoutRenderType,
} from "@pnp/modern-search-extensibility";
import { IWebPartTitleProps } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { IAnalyticsService } from "../../../services/analyticsService/IAnalyticsService";
import ISearchResultsWebPartProps from "../ISearchResultsWebPartProps";

export interface ISearchResultsContainerProps {
   /**
    * The current Web Part data context
    */
   dataContext: IDataContext;

   /**
    * The current page context
    */
   pageContext: PageContext;

   /**
    * The current page context
    */
   teamsContext: any;

   /**
    * The selected data source instance
    */
   dataSource: IDataSource;

   /**
    * The selected data source key
    */
   dataSourceKey: string;

   /**
    * The template content before processing
    */
   templateContent: string;

   /**
    * The Web Part properties so they can be used in Handlebars template
    */
   properties: ISearchResultsWebPartProps;

   /**
    * The current theme information
    */
   themeVariant: IReadonlyTheme;

   /**
    * The Web Part instance ID
    */
   instanceId: string;

   /**
    * Handler when the data have been retrieved from the source. Useful to list all available fields from the source.
    */
   onDataRetrieved: (
      availableDataSourceFields: string[],
      filters?: IDataFilterResult[],
      pageNumber?: number
   ) => void;

   /**
    * Handler when a item has been selected from results
    */
   onItemSelected: (currentSelectedItems: { [key: string]: any }[]) => void;

   /**
    * Handler when no results have been found
    */
   onNoResultsFound: () => void;

   /**
    * The current service scope
    */
   serviceScope: ServiceScope;

   /**
    * The Web Part Title props
    */
   webPartTitleProps: IWebPartTitleProps;

   /**
    * The layout render type (Handlebars, Adaptive Cards, etc.)
    */
   renderType: LayoutRenderType;

   analyticsService: IAnalyticsService;
}
