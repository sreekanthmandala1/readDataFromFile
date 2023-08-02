import { Component } from '@angular/core';
import * as XLSX from "xlsx";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  title = 'readDataFromFile';
  headers: any;
  currentRow: any;
  object: any;
  objectsEl:any;
  email:any;
  name:any;
  objects:any;
  element:any;

  createObjectsFromExcelFile(file: File) {
    debugger
    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    reader.onload = () => {
      const data = new Uint8Array(reader.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
       this.objects = [];
      this.headers = rows[0];
      for (let i = 1; i < rows.length; i++) {
         this.currentRow = rows[i];
        this.object = {};
        for (let j = 0; j < this.headers.length; j++) {
          this.object[this.headers[j]] = this.currentRow[j];
        }
        this.objects.push(this.object);
      }
      console.log(this.objects);
      this.objectsEl = this.objects.forEach((el:any)=>{
        console.log(el)
        this.element = el
        this.email = el.Email
        console.log(this.email)
        this.name = el.Name
      })
    };
  }

  onFileSelected(event:any) {
    debugger
    const file: File = event.target.files[0];
    if (file) {
      this.createObjectsFromExcelFile(file);
    }
  }
}
