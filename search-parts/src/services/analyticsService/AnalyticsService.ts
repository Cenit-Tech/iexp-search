import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import { IAnalyticsItem, IAnalyticsService } from "./IAnalyticsService";

export interface IAnalyticsSession {
   correlationId: string;
   timeAccessed: number;
}

const AnalyticsService_ServiceKey = "IEXPModernSearchAnalyticsService";
const LOCAL_STORAGE_KEY = "iexp-search-analytics";

export class AnalyticsService implements IAnalyticsService {
   public static ServiceKey: ServiceKey<IAnalyticsService> = ServiceKey.create(
      AnalyticsService_ServiceKey,
      AnalyticsService
   );

   public isEnabled: boolean = true;
   private _spSiteUrl: string;
   private _spListName: string;
   private _source: string;
   private _pageContext: PageContext;
   private _spHttpClient: SPHttpClient;

   constructor(serviceScope: ServiceScope) {
      serviceScope.whenFinished(() => {
         this._pageContext = serviceScope.consume<PageContext>(
            PageContext.serviceKey
         );
         this._spHttpClient = serviceScope.consume<SPHttpClient>(
            SPHttpClient.serviceKey
         );
         sp.setup({ spfxContext: { pageContext: this._pageContext } });
      });
   }

   private getStorageItem(key: string): IAnalyticsSession {
      const item = localStorage.getItem(key)
         ? (JSON.parse(localStorage.getItem(key)) as IAnalyticsSession)
         : null;

      if (item) {
         const tenMinsAgo = new Date(Date.now() - 10 * 60 * 1000).getTime();
         let correlationId = item.correlationId;

         // If the item is older than 10 minutes, create a new correlation ID - assume a new session after 10 mins of inactivity
         if (item.timeAccessed < tenMinsAgo) {
            correlationId = crypto.randomUUID();
         }

         // Overwrite the time accessed
         const newItem = {
            correlationId: correlationId,
            timeAccessed: Date.now(),
         };

         localStorage.setItem(key, JSON.stringify(newItem));

         return newItem;
      } else {
         const newItem = {
            correlationId: crypto.randomUUID(),
            timeAccessed: Date.now(),
         };

         localStorage.setItem(key, JSON.stringify(newItem));
         return newItem;
      }
   }

   public init(
      source: string,
      isEnabled: boolean,
      spSiteUrl: string,
      spListName: string
   ): void {
      this.isEnabled = isEnabled;
      this._spListName = spListName;
      this._spSiteUrl = spSiteUrl;
      this._source = source;

      console.debug(
         "Analytics INIT",
         source,
         this.isEnabled,
         this._spSiteUrl,
         this._spListName
      );

      let session = this.getStorageItem(LOCAL_STORAGE_KEY);
   }

   /**
    * Add an event to the analytics service
    * @param event event to add
    * @param properties event item properties
    */
   public async add(event: string, properties: IAnalyticsItem): Promise<void> {
      let session = this.getStorageItem(LOCAL_STORAGE_KEY);
      properties.correlationId = session.correlationId;

      const item = await this._addItemToSharePoint(event, properties);

      console.debug("Analytics ADD", item);
   }

   private async _addItemToSharePoint(
      event: string,
      item: IAnalyticsItem
   ): Promise<any> {
      const web = Web(this._spSiteUrl);
      const list = web.lists.getByTitle(this._spListName);

      const newItem = await list.items.add({
         Title: item.correlationId,
         Source: this._source || "",
         URL: item.URL || this._pageContext.web.absoluteUrl,
         Action: event || "",
         QueryText: item.QueryText || "",
         ResultCount: item.ResultCount,
         Page: item.Page || 0,
         ActionUrl: item.ActionUrl || "",
         ActionValue: item.ActionValue || "",
         AdditionalInfo: item.Properties ? JSON.stringify(item.Properties) : "",
      });

      return newItem;
   }
}
