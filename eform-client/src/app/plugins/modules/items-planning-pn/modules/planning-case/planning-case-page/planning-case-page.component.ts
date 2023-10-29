import {
  Component,
  OnInit,
  QueryList,
  ViewChild,
  ViewChildren,
} from '@angular/core';
import { ActivatedRoute, Router } from '@angular/router';
import { EFormService, CasesService } from 'src/app/common/services';
import {
  TemplateDto,
  CaseEditRequest,
  ReplyElementDto,
  ReplyRequest, ElementDto, DataItemDto,
} from 'src/app/common/models';
import { AuthStateService } from 'src/app/common/store';
import { CaseEditElementComponent } from 'src/app/common/modules/eform-cases/components';
import {ItemsPlanningPnCasesService} from '../../../services';
import {DateTimeAdapter} from '@danielmoncada/angular-datetime-picker';
import * as R from 'ramda';
import {selectCurrentUserLocale} from 'src/app/state/auth/auth.selector';
import {Store} from '@ngrx/store';

@Component({
  selector: 'app-installation-case-page',
  templateUrl: './planning-case-page.component.html',
  styleUrls: ['./planning-case-page.component.scss'],
})
export class PlanningCasePageComponent implements OnInit {
  @ViewChildren(CaseEditElementComponent)
  editElements: QueryList<CaseEditElementComponent>;
  @ViewChild('caseConfirmation', { static: false }) caseConfirmation;
  id: number;
  planningId: number;
  eFormId: number;
  dateFrom: string;
  dateTo: string;
  currentTemplate: TemplateDto = new TemplateDto();
  replyElement: ReplyElementDto = new ReplyElementDto();
  reverseRoute: string;
  requestModels: Array<CaseEditRequest> = [];
  replyRequest: ReplyRequest = new ReplyRequest();
  maxDate: Date;
  initialDate: Date;
  private selectCurrentUserLocale$ = this.authStore.select(selectCurrentUserLocale);

  constructor(
    dateTimeAdapter: DateTimeAdapter<any>,
    private activateRoute: ActivatedRoute,
    private authStore: Store,
    private casesService: CasesService,
    private eFormService: EFormService,
    private router: Router,
    private authStateService: AuthStateService,
    private itemsPlanningPnCasesService: ItemsPlanningPnCasesService
  ) {
    this.activateRoute.params.subscribe((params) => {
      this.id = +params['id'];
      this.planningId = +params['planningId'];
      this.eFormId = +params['templateId'];
      this.dateFrom = params['dateFrom'];
      this.dateTo = params['dateTo'];
      this.selectCurrentUserLocale$.subscribe((locale) => {
        dateTimeAdapter.setLocale(locale);
      });
    });
  }

  ngOnInit() {
    this.loadTemplateInfo();
    this.maxDate = new Date();
  }

  loadCase() {
    if (!this.id || this.id === 0) {
      return;
    }
    this.casesService
      .getById(this.id, this.currentTemplate.id)
      .subscribe((operation) => {
        if (operation && operation.success) {
          this.replyElement = operation.model;
          this.initialDate = this.replyElement.doneAt;
        }
      });
  }

  loadTemplateInfo() {
    if (this.eFormId) {
      this.eFormService.getSingle(this.eFormId).subscribe((operation) => {
        if (operation && operation.success) {
          this.currentTemplate = operation.model;
          this.loadCase();
        }
      });
    }
  }

  saveCase(navigateToPosts?: boolean) {
    this.requestModels = [];
    this.editElements.forEach((x) => {
      x.extractData();
      this.requestModels.push(x.requestModel);
    });
    this.replyRequest.id = this.replyElement.id;
    this.replyRequest.label = this.replyElement.label;
    this.replyRequest.elementList = this.requestModels;
    if (this.initialDate !== this.replyElement.doneAt) {
      this.replyRequest.doneAt =
        new Date(
          Date.UTC(
            this.replyElement.doneAt.getFullYear(),
            this.replyElement.doneAt.getMonth(),
            this.replyElement.doneAt.getDate(),
            0,
            0,
            0
          )
        );
    } else {
      this.replyRequest.doneAt = this.replyElement.doneAt;
    }
    this.itemsPlanningPnCasesService
      .updateCase(this.replyRequest, this.currentTemplate.id)
      .subscribe((operation) => {
        if (operation && operation.success) {
          this.replyElement = new ReplyElementDto();
          this.router
            .navigate([
              '/plugins/items-planning-pn/reports/' +
                this.dateFrom +
                '/' +
                this.dateTo,
            ])
            .then();
        }
      });
  }

  goToSection(location: string): void {
    window.location.hash = location;
    setTimeout(() => {
      document.querySelector(location).parentElement.scrollIntoView();
    });
  }

  partialLoadCase() {
    if (!this.id || this.id === 0) {
      return;
    }
    this.casesService
      .getById(this.id, this.currentTemplate.id)
      .subscribe((operation) => {
        if (operation && operation.success) {
          const fn = (pathForLens: Array<number | string>) => {
            const lens = R.lensPath(pathForLens);
            let dataItem: (ElementDto | DataItemDto) = R.view(lens, operation.model);
            // @ts-ignore
            if (dataItem.elementList !== undefined || dataItem.dataItemList !== undefined) {
              dataItem = dataItem as ElementDto;
              // R.set(R.lensPath([...pathForLens, 'extraPictures']), dataItem.extraPictures, this.replyElement);
              if(dataItem.elementList) {
                for (let i = 0; i < dataItem.elementList.length; i++) {
                  fn([...pathForLens, 'elementList', i]);
                }
              }
              if(dataItem.dataItemList) {
                for (let i = 0; i < dataItem.dataItemList.length; i++) {
                  fn([...pathForLens, 'dataItemList', i]);
                }
              }
            } else { // @ts-ignore
              if(dataItem.fieldType !== undefined){
                dataItem = dataItem as DataItemDto;
                if (dataItem.fieldType === 'FieldContainer') {
                  for (let i = 0; i < dataItem.dataItemList.length; i++) {
                    fn([...pathForLens, 'dataItemList', i]);
                  }
                }
                if (dataItem.fieldType === 'Picture') {
                  // let oldDataItem = R.view(lens, this.replyElement);
                  // oldDataItem = {...oldDataItem, ...dataItem};
                  this.replyElement = R.set(lens, dataItem, this.replyElement);
                }
              }
            }
          }
          for (let i = 0; i < operation.model.elementList.length; i++){
            fn(['elementList', i]);
          }
        }
      });
  }
}
