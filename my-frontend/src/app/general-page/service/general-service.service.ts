import { Injectable } from '@angular/core';
import {HttpClient} from "@angular/common/http";
import {Observable} from "rxjs";

@Injectable({
  providedIn: 'root'
})
export class GeneralServiceService {
private url='http://localhost:8080/api/lab'
  constructor(private http:HttpClient) { }

  getTest():Observable<any>{
  return this.http.get(`${this.url}`)
  }
}
