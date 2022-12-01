import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input, OnChanges,
  OnInit,
  Output, SimpleChanges,
  ViewChild,
} from '@angular/core';
import {ReportEformItemModel} from '../../../../models';
import {AuthStateService} from 'src/app/common/store';
import {ViewportScroller} from '@angular/common';
import {Router} from '@angular/router';
import {PlanningsReportStateService} from './../store';
import {MtxGridColumn} from '@ng-matero/extensions/grid';
import {TranslateService} from '@ngx-translate/core';
import {MatDialog} from '@angular/material/dialog';
import {Overlay} from '@angular/cdk/overlay';
import {dialogConfigHelper} from 'src/app/common/helpers';
import {PlanningCaseDeleteComponent} from 'src/app/plugins/modules/items-planning-pn/modules/reports/components';

@Component({
  selector: 'app-report-table',
  templateUrl: './report-table.component.html',
  styleUrls: ['./report-table.component.scss'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ReportTableComponent implements OnInit, OnChanges {
  @Input() items: ReportEformItemModel[] = [];
  @Input() reportIndex: number;
  @Input() dateFrom: any;
  @Input() dateTo: any;
  @Input() itemHeaders: { key: string; value: string }[] = [];
  @Input() newPostModal: any;
  @Output() planningCaseDeleted: EventEmitter<void> = new EventEmitter<void>();
  @Output() btnViewPicturesClicked: EventEmitter<{ reportIndex: number, caseId: number }>
    = new EventEmitter<{ reportIndex: number, caseId: number }>();

  tableHeaders: MtxGridColumn[] = [
    {header: this.translateService.stream('Id'), field: 'microtingSdkCaseId'},
    {header: this.translateService.stream('CreatedAt'), field: 'microtingSdkCaseDoneAt', type: 'date', typeParameter: {format: 'dd.MM.y'}},
    {header: this.translateService.stream('Done by'), field: 'doneBy',},
    {header: this.translateService.stream('Item name'), field: 'itemName',},
    {
      header: this.translateService.stream('Actions'),
      field: 'actions',
      type: 'button',
      buttons: [
        {
          tooltip: this.translateService.stream('View images'),
          type: 'icon',
          click: (record: ReportEformItemModel) => this.onClickViewPicture(record.microtingSdkCaseId),
          color: 'accent',
          icon: 'image',
          iif: (record: ReportEformItemModel) => record.imagesCount !== 0,
        },
        {
          tooltip: this.translateService.stream('Edit'),
          type: 'icon',
          click: (record: ReportEformItemModel) => this.onClickEditCase(record.microtingSdkCaseId, record.eFormId, record.id),
          color: 'accent',
          icon: 'edit',
        },
        {
          tooltip: this.translateService.stream('Delete'),
          type: 'icon',
          click: (record: ReportEformItemModel) => this.onShowDeletePlanningCaseModal(record),
          color: 'warn',
          icon: 'delete',
        }
      ]
    },
  ];
  mergedTableHeaders: MtxGridColumn[] = [];
  caseDeleteComponentComponentAfterClosedSub$: any;

  constructor(private authStateService: AuthStateService,
              private viewportScroller: ViewportScroller,
              private router: Router,
              private planningsReportStateService: PlanningsReportStateService,
              private translateService: TranslateService,
              private dialog: MatDialog,
              private overlay: Overlay,
  ) {
  }

  ngOnInit(): void {
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (/*!changes.itemHeaders.isFirstChange() && */changes.itemHeaders) {
      let itemHeaders = this.itemHeaders.map((x, i): MtxGridColumn => {
        return {
          header: x.value,
          field: x.value,
          formatter: (record: ReportEformItemModel) => {
            if (record.caseFields[i].value === 'checked') {
              return `<span class="material-icons">done</span>`;
            }
            if (record.caseFields[i].value !== 'checked' && record.caseFields[i].value !== 'unchecked') {
              // @ts-ignore
              if (record.caseFields[i].key === 'number') {
                return record.caseFields[i].value.replace('.', ',');
              }
              // @ts-ignore
              if (record.caseFields[i].key !== 'number') {
                return record.caseFields[i].value;
              }
            }
          },
        };
      });
      this.mergedTableHeaders = [...this.tableHeaders, ...itemHeaders];
    }
  }

  // openCreateModal(
  //   caseId: number,
  //   eformId: number,
  //   pdfReportAvailable: boolean
  // ) {
  //   this.newPostModal.caseId = caseId;
  //   this.newPostModal.efmroId = eformId;
  //   this.newPostModal.currentUserFullName = this.authStateService.currentUserFullName;
  //   this.newPostModal.pdfReportAvailable = pdfReportAvailable;
  //   this.newPostModal.show();
  // }

  onShowDeletePlanningCaseModal(item: ReportEformItemModel) {
    this.caseDeleteComponentComponentAfterClosedSub$ = this.dialog.open(PlanningCaseDeleteComponent,
      {...dialogConfigHelper(this.overlay, item)})
      .afterClosed().subscribe(data => data ? this.onPlanningCaseDeleted() : undefined);
  }

  onPlanningCaseDeleted() {
    this.planningCaseDeleted.emit();
  }

  onClickViewPicture(caseId: number) {
    this.btnViewPicturesClicked.emit({reportIndex: this.reportIndex, caseId});
  }

  onClickEditCase(microtingSdkCaseId: number, eFormId: number, id: number) {
    this.planningsReportStateService.updateScrollPosition(this.viewportScroller.getScrollPosition());
    this.router.navigate([`/plugins/items-planning-pn/case/`, microtingSdkCaseId, eFormId, id, this.dateFrom, this.dateTo])
      .then();
  }
}
