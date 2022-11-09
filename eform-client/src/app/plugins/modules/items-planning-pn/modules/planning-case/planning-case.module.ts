import {NgModule} from '@angular/core';
import {CommonModule} from '@angular/common';
import {TranslateModule} from '@ngx-translate/core';
import {MDBBootstrapModule} from 'angular-bootstrap-md';
import {NgSelectModule} from '@ng-select/ng-select';
import {GallerizeModule} from '@ngx-gallery/gallerize';
import {LightboxModule} from '@ngx-gallery/lightbox';
import {GalleryModule} from '@ngx-gallery/core';
import {FormsModule} from '@angular/forms';
import {FontAwesomeModule} from '@fortawesome/angular-fontawesome';
import {PlanningCaseRouting} from './planning-case.routing';
import {CasesModule} from 'src/app/modules';
import {EformImportedModule} from 'src/app/common/modules/eform-imported/eform-imported.module';
import {EformSharedModule} from 'src/app/common/modules/eform-shared/eform-shared.module';
import {PlanningCasePageComponent} from './planning-case-page/planning-case-page.component';
import {PlanningCaseHeaderComponent} from './planning-case-header/planning-case-header.component';
import {EformCasesModule} from 'src/app/common/modules/eform-cases/eform-cases.module';
import {OwlDateTimeModule} from '@danielmoncada/angular-datetime-picker';

@NgModule({
  declarations: [
    PlanningCaseHeaderComponent,
    PlanningCasePageComponent
  ],
  imports: [
    TranslateModule,
    MDBBootstrapModule,
    EformSharedModule,
    PlanningCaseRouting,
    CommonModule,
    NgSelectModule,
    EformImportedModule,
    GallerizeModule,
    LightboxModule,
    GalleryModule,
    FormsModule,
    FontAwesomeModule,
    CasesModule,
    EformCasesModule,
    OwlDateTimeModule
  ]
})
export class PlanningCaseModule {
}
