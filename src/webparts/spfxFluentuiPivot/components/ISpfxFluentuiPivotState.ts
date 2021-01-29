type user = {
  Title: string;
}

type group = {
  Title: string;
}

export interface ISpfxFluentuiPivotState {
  siteUsres: user[];
  siteGroups: group[];
}
