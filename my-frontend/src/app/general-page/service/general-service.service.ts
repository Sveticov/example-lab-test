import { Injectable } from '@angular/core';
import {HttpClient, HttpEvent, HttpRequest} from "@angular/common/http";
import {Observable} from "rxjs";
import {FileData} from "../file-data";

@Injectable({
  providedIn: 'root'
})
export class GeneralServiceService {
private url='http://localhost:8080/api/lab'
  constructor(private http:HttpClient) { }

  getTest():Observable<any>{
  return this.http.get(`${this.url}`)
  }

  sendFile(file:FileData):Observable<any>{
  return this.http.post<FileData>(this.url+"/filet",file)
  }

  upload(currentFile: File | undefined) :Observable<HttpEvent<any>>{
    const formData = new FormData();

    // @ts-ignore
    formData.append('file',currentFile);
    const  req = new HttpRequest('POST',`${this.url}/files`,formData,{
      reportProgress:true,
      responseType:'json'
    });
    return  this.http.request(req)
  }
}
