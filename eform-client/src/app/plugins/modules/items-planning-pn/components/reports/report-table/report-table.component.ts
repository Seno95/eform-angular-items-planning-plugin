import {
  ChangeDetectionStrategy,
  Component,
  EventEmitter,
  Input,
  OnInit,
  Output,
  ViewChild,
} from '@angular/core';
import {ReportEformItemModel } from '../../../models/report';
import {AuthStateService} from 'src/app/common/store';

@Component({
  selector: 'app-report-table',
  templateUrl: './report-table.component.html',
  styleUrls: ['./report-table.component.scss'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class ReportTableComponent implements OnInit {
  @Input() items: ReportEformItemModel[] = [];
  @Input() reportIndex: number;
  @Input() dateFrom: any;
  @Input() dateTo: any;
  @Input() itemHeaders: { key: string; value: string }[] = [];
  @Input() newPostModal: any;
  @ViewChild('deletePlanningCaseModal') deletePlanningCaseModal;
  @Output() planningCaseDeleted: EventEmitter<void> = new EventEmitter<void>();
  @Output() btnViewPicturesClicked: EventEmitter<{ reportIndex: number, caseId: number }>
    = new EventEmitter<{ reportIndex: number, caseId: number }>();

  constructor(private authStateService: AuthStateService) {
  }

  ngOnInit(): void {
  }

  openCreateModal(
    caseId: number,
    eformId: number,
    pdfReportAvailable: boolean
  ) {
    this.newPostModal.caseId = caseId;
    this.newPostModal.efmroId = eformId;
    this.newPostModal.currentUserFullName = this.authStateService.currentUserFullName;
    this.newPostModal.pdfReportAvailable = pdfReportAvailable;
    this.newPostModal.show();
  }

  onShowDeletePlanningCaseModal(item: ReportEformItemModel) {
    this.deletePlanningCaseModal.show(item);
  }

  onPlanningCaseDeleted() {
    this.planningCaseDeleted.emit();
  }

  onClickViewPicture(caseId: number) {
    this.btnViewPicturesClicked.emit({reportIndex: this.reportIndex, caseId});
  }
}
