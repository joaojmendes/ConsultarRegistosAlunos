export default interface IListViewData {
  key: string;
  Title: string;  // nome colaborador
  nmecanografico: string;
  nifcolaborador: string;
  nomealuno: string;
  DataNascAluno: string;
  IdadeAluno:string
  NIFAluno: string;
  Holding: string;
  area:string;
  codigopostal: string;
  localidade: string;
  media:string;
  morada: string;
  empresa: string;
  nomeloja: string;
  consentimento: string | any;
  ano: string;
  attachements?:any[],
}
