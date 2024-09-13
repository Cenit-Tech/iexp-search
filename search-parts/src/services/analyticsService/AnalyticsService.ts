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
   sessionId: string;
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
   public enabledOnlyWithQueryText: boolean = true;
   private _spSiteUrl: string;
   private _spListName: string;
   private _source: string;
   private _pageContext: PageContext;
   private _spHttpClient: SPHttpClient;
   private _pageNumber: number = 0;
   private _QueryId: string;
   private _queryText: string;

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
      const item = sessionStorage.getItem(key)
         ? (JSON.parse(sessionStorage.getItem(key)) as IAnalyticsSession)
         : null;

      if (item) {
         const tenMinsAgo = new Date(Date.now() - 10 * 60 * 1000).getTime();
         let sessionId = item.sessionId;

         // If the item is older than 10 minutes, create a new correlation ID - assume a new session after 10 mins of inactivity
         if (item.timeAccessed < tenMinsAgo) {
            sessionId = crypto.randomUUID();
         }

         // Overwrite the time accessed
         const newItem = {
            sessionId: sessionId,
            timeAccessed: Date.now(),
         };

         sessionStorage.setItem(key, JSON.stringify(newItem));

         return newItem;
      } else {
         const newItem = {
            sessionId: crypto.randomUUID(),
            timeAccessed: Date.now(),
         };

         sessionStorage.setItem(key, JSON.stringify(newItem));
         return newItem;
      }
   }

   public init(
      source: string,
      isEnabled: boolean,
      enabledOnlyWithQueryText: boolean,
      spSiteUrl: string,
      spListName: string
   ): void {
      this.isEnabled = isEnabled;
      this.enabledOnlyWithQueryText = enabledOnlyWithQueryText;
      this._spListName = spListName;
      this._spSiteUrl = spSiteUrl;
      this._source = source;

      console.debug(
         "Search Analytics Init",
         source,
         this.isEnabled,
         this._spSiteUrl,
         this._spListName
      );
   }

   /**
    * Add an event to the analytics service
    * @param event event to add
    * @param properties event item properties
    */
   public async add(event: string, properties: IAnalyticsItem): Promise<void> {
      if (!this.isEnabled) {
         return;
      }
      if (this.enabledOnlyWithQueryText && !properties.QueryText) {
         console.debug("SKIP Search Analytics - No QueryText");
         return;
      }

      if (this._queryText != properties.QueryText) {
         // start new search group
         this._queryText = properties.QueryText;
         this._QueryId = crypto.randomUUID();
      }

      properties.QueryId = this._QueryId;

      let session = this.getStorageItem(LOCAL_STORAGE_KEY);
      properties.sessionId = session.sessionId;

      if (properties.Page) {
         this._pageNumber = properties.Page;
      }

      const item = await this._addItemToSharePoint(event, properties);
   }

   public addResultHooks(HtmlDivElement: HTMLElement) {
      const links = HtmlDivElement.querySelectorAll("a");

      for (let i = 0; i < links.length; i++) {
         const link = links[i];

         // Ignore specific links
         if (link.classList.contains("ignore")) {
            continue;
         }

         const url = link.getAttribute("href");
         link.setAttribute("data-href", url);
         link.setAttribute("data-index", (i + 1).toString());
         link.setAttribute("href", "#");

         link.addEventListener("click", (e) => {
            e.preventDefault();

            console.log(url);

            let session = this.getStorageItem(LOCAL_STORAGE_KEY);

            this._addItemToSharePoint("ResultClick", {
               sessionId: session.sessionId,
               QueryId: this._QueryId,
               ActionUrl: url,
               Action: "ResultClick",
               Source: this._source || "",
               ActionValue: link.getAttribute("data-index"),
               Page: this._pageNumber,
            });

            location.href = url;
         });
      }
   }

   private async _addItemToSharePoint(
      event: string,
      item: IAnalyticsItem
   ): Promise<any> {
      if (!this.isEnabled) {
         return;
      }

      const web = Web(this._spSiteUrl);
      const list = web.lists.getByTitle(this._spListName);

      const newItem = await list.items.add({
         Title: item.sessionId,
         QueryId: item.QueryId,
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

      console.debug("Analytics ADD", newItem);

      return newItem;
   }
}
