export interface ISoapTaxonomyResponse {
    description: string;
    labels: { value: string; isDefault: boolean }[];
    text:string;
    info?: {
      parentId: string;
      parentLabel: string;
      termPath: string;
      children: string[];
      hasChildren: boolean;
      termSetId: string;
      termSetLabel: string;
    }[];
    id: string;
    isDeprecated: boolean;
    internalId: string;
    hasChildren: boolean;
    selected:boolean;
    children: ISoapTaxonomyResponse[];
  }