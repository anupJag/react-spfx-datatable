export interface IListColumnOptions{
    key : string;
    text : string;
}

export interface IUserProfile{
    Title : string;
}

export interface IHyperLink{
    Description : string;
    Url : string;
}

export interface IDataCache{
    expirationTime : string;
    results : any[];
}

export enum FieldType {
    Integer = 1,
    SingleLineText = 2,
    MultiLineText = 3,
    DateTime = 4,
    Choice = 6,
    LookUp = 7,
    Boolean = 8,
    Number = 9,
    Currency = 10,
    URL = 11,
    MultiChoice = 15,
    People = 20,
}

export enum FieldTypeNames {
    Integer = "Integer",
    SingleLineText = "SingleLineText",
    MultiLineText = "MultiLineText",
    DateTime = "DateTime",
    Choice = "Choice",
    LookUp = "LookUp",
    Boolean = "Boolean",
    Number = "Number",
    Currency = "Currency",
    URL = "URL",
    MultiChoice = "MultiChoice",
    People = "People",
}

export type MapSchemaTypes = {
    Integer : number;
    SingleLineText : string;
    MultiLineText : string;
    DateTime : string;
    Choice : string;
    LookUp : string;
    Boolean : boolean;
    Number : number;
    Currency : string;
    URL : IHyperLink[];
    MultiChoice : string[];
    People : IUserProfile[];
};

export type MapSchema<T extends Record<string, keyof MapSchemaTypes>> = {
    [K in keyof T] : MapSchemaTypes[T[K]]
};


export interface QueryStructure {
    queryParameter : string[];
    expandParameter : string[];
}

