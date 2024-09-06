export interface IAnalyticsService {
   isEnabled: boolean;
   init(
      source: string,
      isEnabled: boolean,
      spSiteUrl: string,
      spListName: string
   ): void;
   add(event: string, properties: IAnalyticsItem): void;
}

export interface IAnalyticsItem {
   correlationId?: string;
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
