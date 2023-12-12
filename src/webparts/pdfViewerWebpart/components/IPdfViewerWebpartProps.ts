import { PdfViewerEntity } from "../../../sp/entities/pdf-viewer-webpart";
import { IDataAccessService } from "../../../sp/services/data-access-service";

export interface IPdfViewerWebpartProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  dataService: IDataAccessService;
}

export interface IPdfDataAccessState
{
    listitem: PdfViewerEntity;

    allItems: PdfViewerEntity[];

    saving: boolean;
}