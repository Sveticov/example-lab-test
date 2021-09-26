import {Component, OnInit} from '@angular/core';
import {GeneralServiceService} from "./service/general-service.service";
import {NgForm} from "@angular/forms";
import {FileData} from "./file-data";
import {HttpEventType, HttpResponse} from "@angular/common/http";

@Component({
  selector: 'app-general-page',
  templateUrl: './general-page.component.html',
  styleUrls: ['./general-page.component.css']
})
export class GeneralPageComponent implements OnInit {
  selectedFiles?: FileList;
  currentFile?: File;
  message = ''
  errorMessage = ''
  progress = 0
  visiNextPage = false


  constructor(private service: GeneralServiceService) {
  }

  ngOnInit(): void {
  }

  getTest() {
    this.service.getTest().subscribe(data => console.log(data))
  }

  sendFileData(myFile: NgForm) {
    this.service.sendFile(new FileData(myFile.value.name, myFile.value.size)).subscribe(data => console.log(data))
  }

  selectFile(event: any) {
    this.selectedFiles = event.target.files
    this.progress=0
    console.log(event.target.files)
  }

  upload(): void {
    this.errorMessage = '';
    this.progress=0

    if (this.selectedFiles) {
      const file: File | null = this.selectedFiles.item(0)
      if (file) this.currentFile = file
      this.service.upload(this.currentFile).subscribe(
        (data: any) => {
          if (data.type === HttpEventType.UploadProgress) {
            this.progress=(Math.round(100 * data.loaded / data.total));
          } else if (data instanceof HttpResponse) {
            this.message = data.body.responseMessage;
          }
        },
        (err: any) => {
          console.log(err);
          if (err.error && err.error.responseMessage) {
            this.errorMessage = err.error.responseMessage;
          } else {
            this.errorMessage = 'Error file'
          }
          this.currentFile = undefined
        });
    }
    this.selectedFiles = undefined
  }


  nextPage() {
    this.visiNextPage = !this.visiNextPage
  }
}
