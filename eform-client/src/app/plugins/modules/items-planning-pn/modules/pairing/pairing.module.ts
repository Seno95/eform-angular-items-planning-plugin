import {CommonModule} from '@angular/common';
import {NgModule} from '@angular/core';
import {RouterModule} from '@angular/router';
import {
  OwlDateTimeModule,
  OwlMomentDateTimeModule,
} from '@danielmoncada/angular-datetime-picker';
import {TranslateModule} from '@ngx-translate/core';
import {
  PairingGridPageComponent,
  PairingGridTableComponent,
  PairingGridUpdateComponent,
} from './components';
import {PairingRouting} from './pairing.routing';
import {EformSharedModule} from 'src/app/common/modules/eform-shared/eform-shared.module';
import {FontAwesomeModule} from '@fortawesome/angular-fontawesome';
import {MtxGridModule} from '@ng-matero/extensions/grid';
import {FormsModule, ReactiveFormsModule} from '@angular/forms';
import {EformCasesModule} from 'src/app/common/modules/eform-cases/eform-cases.module';
import {EformSharedTagsModule} from 'src/app/common/modules/eform-shared-tags/eform-shared-tags.module';
import {pairingPersistProvider} from './components/store';
import {MatButtonModule} from '@angular/material/button';
import {MatTooltipModule} from '@angular/material/tooltip';
import {MatFormFieldModule} from '@angular/material/form-field';
import {MtxSelectModule} from '@ng-matero/extensions/select';
import {MatCheckboxModule} from '@angular/material/checkbox';
import {MatDialogModule} from '@angular/material/dialog';

@NgModule({
  declarations: [
    PairingGridPageComponent,
    PairingGridTableComponent,
    PairingGridUpdateComponent,
  ],
    imports: [
        CommonModule,
        TranslateModule,
        RouterModule,
        OwlDateTimeModule,
        PairingRouting,
        OwlMomentDateTimeModule,
        EformSharedModule,
        FontAwesomeModule,
        MtxGridModule,
        FormsModule,
        ReactiveFormsModule,
        EformCasesModule,
        EformSharedTagsModule,
        MatButtonModule,
        MatTooltipModule,
        MatFormFieldModule,
        MtxSelectModule,
        MatCheckboxModule,
        MatDialogModule,
    ],
  providers: [pairingPersistProvider],
})
export class PairingModule {
}