export interface ITaxonomyElementTL {
    _a31: string;
    _a32: string;
}

export interface ITaxonomyElementTD {
    a11: string;
}

export interface ITaxonomyElementTM {
    _a12: string;
    _a17: string;
    _a24: string;
    _a25: string;
    _a40: string;
    _a45: string;
    _a67: string;
    _a69: string;
}

export interface ITaxonomyElementTMS {
    TM: ITaxonomyElementTM;
}


export interface ITaxonomyElementLS {
    TL: ITaxonomyElementTL;
}


export interface ITaxonomyElementT {
    DS: string;
    TD: ITaxonomyElementTD;
    LS: ITaxonomyElementLS;
    TMS: ITaxonomyElementTMS;
    _a9: string;
    _a21: string;
    _a61: string;
    _a1000: string;
}

export interface ITermStore {
    T: ITaxonomyElementT[];
}

export interface ITaxonomyResponse {
    TermStore: ITermStore;
}