export interface IFile {
  Id: number;
  Title: string;
  Name: string;
  Size: number;
}

// create PnP JS response interface for File
export interface IResponseFile {
  Length: number;
}

// create PnP JS response interface for Item
export interface IResponseItem {
  Id: number;
  File: IResponseFile;
  FileLeafRef: string;
  Title: string;
}

// Vicky Student List interface
export interface IVickyStudentList {
  Title: number;
  StudName: string;
  StudDept: string;
  StudCity: string;
}
