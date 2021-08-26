import { Component, OnInit } from '@angular/core';
import {GeneralServiceService} from "./service/general-service.service";

@Component({
  selector: 'app-general-page',
  templateUrl: './general-page.component.html',
  styleUrls: ['./general-page.component.css']
})
export class GeneralPageComponent implements OnInit {

  constructor(private service:GeneralServiceService) { }

  ngOnInit(): void {
  }

  getTest() {
    this.service.getTest().subscribe(data=>console.log(data))
  }
}
