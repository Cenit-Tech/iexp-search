export interface IAnalyticsService {
   isEnabled: boolean;
   init(
      source: string,
      isEnabled: boolean,
      enabledOnlyWithQueryText: boolean,
      spSiteUrl: string,
      spListName: string
   ): void;
   add(event: string, properties: IAnalyticsItem): void;
   addResultHooks(HtmlDivElement: HTMLElement): void;
}

export interface IAnalyticsItem {
   sessionId?: string;
   QueryId?: string;
   Source?: string;
   URL?: string;
   Action?: string;
   QueryText?: string;
   ResultCount?: number;
   Page?: number;
   ActionUrl?: string;
   ActionValue?: string;
   Properties?: any;
}
