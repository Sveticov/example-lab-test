import { Injectable } from '@angular/core';
import {Observable} from "rxjs";
import {HttpClient, HttpEvent, HttpRequest} from "@angular/common/http";

@Injectable({
  providedIn: 'root'
})
export class SecondServiceService {
  private url = 'http://localhost:8080/api/lab'

  constructor(private http: HttpClient) {
  }


  uploadPlc(currentFilePLC: File | undefined): Observable<HttpEvent<any>> {
    const formDataPlc = new FormData()
    // @ts-ignore
    formDataPlc.append('file_plc', currentFilePLC);
    const req = new HttpRequest('POST', `${this.url}/files_plc`, formDataPlc, {
      reportProgress: true,
      responseType: 'json'
    });
    return this.http.request(req)
  }

  makeTableLaboratoryReport(): Observable<any> {
    return this.http.get(`${this.url}/table`)
  }

}
