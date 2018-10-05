export default interface IListData {
  Id:string;
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
  consentimento: { results: []};
  ano: string;
}
