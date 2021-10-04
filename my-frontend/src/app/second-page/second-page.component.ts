import {Component, OnInit} from '@angular/core';
import {SecondServiceService} from "./service/second-service.service";
import {HttpEventType, HttpResponse} from "@angular/common/http";

@Component({
  selector: 'app-second-page',
  templateUrl: './second-page.component.html',
  styleUrls: ['./second-page.component.css']
})
export class SecondPageComponent implements OnInit {
  selectedFilesPlc?: FileList;
  currentFilePlc?: File;
  messagePlc = ''
  errorMessagePlc = ''
  progressPlc = 0

  constructor(private service: SecondServiceService) {
  }

  ngOnInit(): void {
  }

  selectFilePlc(event: any) {
    this.selectedFilesPlc = event.target.files
    this.progressPlc = 0
    console.log(event.target.files)
  }

  uploadPlc(): void {
    this.errorMessagePlc = ''
    this.progressPlc = 0

    if (this.selectedFilesPlc) {
      const file: File | null = this.selectedFilesPlc.item(0)
      if (file) this.currentFilePlc = file
      this.service.uploadPlc(this.currentFilePlc).subscribe(
        (data: any) => {
          if (data.type === HttpEventType.UploadProgress) {
            this.progressPlc = (Math.round(100 * data.loaded / data.total));
          } else if (data instanceof HttpResponse) {
            this.messagePlc = data.body.responseMessage;
          }
        },
        (err: any) => {
          console.log(err);
          if (err.error && err.error.responseMessage) {
            this.errorMessagePlc = err.error.responseMessage;
          } else {
            this.errorMessagePlc = 'Error file plc'
          }
          this.currentFilePlc = undefined
        });
    }
    this.selectedFilesPlc = undefined
  }

  makeTableLaboratoryReport() {
    this.service.makeTableLaboratoryReport()
      .subscribe(data=>console.log(data))
    setTimeout(()=>{
      window.location.reload();
    },5000);
  }

}
