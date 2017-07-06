export class PeopleSearchQuery {
    public queryParams: PeopleSearchQueryParams;

    /**
     * Default constructor
     */
    constructor(params: PeopleSearchQueryParams) {
        this.queryParams = params;
    }
}

export class PeopleSearchQueryParams {
    public QueryString: string;
    public MaximumEntitySuggestions: number;
    public AllowEmailAddresses: boolean;
    public AllowOnlyEmailAddresses: boolean;
    public PrincipalType: number;
    public PrincipalSource: number;
    public SharePointGroupID: number;
}

export interface IPeopleSearchQueryResult {
    value: string;
}

export interface IPeopleSearchQueryResultItem {
    Key : string;
    Description: string;
    DisplayText: string;
    EntityType: string;
    ProviderDisplayName: string;
    ProviderName: string;
    IsResolved: boolean;
    EntityData: IPeopleSearchQueryResultItemEntityData;
}

export interface IPeopleSearchQueryResultItemEntityData {
    IsAltSecIdPresent: boolean;
    Title: string;
    Email: string;
    MobilePhone: string;
    ObjectId: string;
    Department: string;
    MultipleMatches: boolean;
}