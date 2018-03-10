import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';


import { AppComponent } from './components/app.component';
import { ImportComponent } from './components/import.component';
import { ExportComponent } from './components/export.component';


@NgModule({
  declarations: [
    AppComponent,
    ImportComponent,
    ExportComponent
  ],
  imports: [
    BrowserModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
