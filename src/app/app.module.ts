import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { ReactiveFormsModule } from '@angular/forms';

import { AppComponent } from './app.component';
import { CreditCardDirective } from './credit-card.directive';

@NgModule({
  imports:      [ BrowserModule, ReactiveFormsModule ],
  declarations: [ AppComponent,  CreditCardDirective ],
  bootstrap:    [ AppComponent ]
})
export class AppModule { }
