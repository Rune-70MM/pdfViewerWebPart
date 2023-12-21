/* eslint-disable @typescript-eslint/typedef */
/* eslint-disable @microsoft/spfx/no-async-await */
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PdfViewerEntity } from "../entities/pdf-viewer-webpart";
import { PageContext } from '@microsoft/sp-page-context'
import { SPFI, SPFx, spfi } from '@pnp/sp';
import { IItemAddResult } from "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import Lists from "../constants/list-names";
import FieldName from "../constants/list-field-names";

export interface IDataAccessService
{
    getAll(): Promise<Array<PdfViewerEntity>>;

    getById(listItemId: number): Promise<PdfViewerEntity>;

    addItem(entity: PdfViewerEntity): Promise<number>;

    updateItem(entity: PdfViewerEntity): Promise<boolean>;
}

export class DataAccessService implements IDataAccessService
{

    public static readonly serviceKey: ServiceKey<IDataAccessService> = ServiceKey.create('MF:DataAccessService', DataAccessService)

    private _sp: SPFI;

    public constructor(serviceScope: ServiceScope)
    {
        serviceScope.whenFinished(() =>
        {
            const pageContext = serviceScope.consume(PageContext.serviceKey);

            this._sp = spfi().using(SPFx({ pageContext }));

        })
    }

    public async getAll(): Promise<PdfViewerEntity[]> 
    {
        try 
        {
            let result = new Array<PdfViewerEntity>();

            //const items = await this._sp.web.lists.getById("8A423C53-90E0-4071-B202-9DD1296290B9").items();
            const items = await this._sp.web.lists.getByTitle(Lists.Invoices.Name).items();

            console.log('List Items;', items);

            for (let index = 0; index < items.length; index++)
            {
                const item = items[index];

                const entity = this.mapToListItem(item);

                result.push(entity);
            }


            debugger;

            return result;
        } catch (error)
        {

            debugger;
        }
    }
    public async getById(listItemId: number): Promise<PdfViewerEntity> 
    {
        throw new Error("Method not implemented.");
    }

    private mapToListItem(entity: PdfViewerEntity): any
    {
        try
        {
            const properties: any = {};

            properties[FieldName.Invoices.Title] = entity.Title;
            properties[FieldName.Invoices.Id] = entity.Id;
            properties[FieldName.Invoices.Owner] = entity.Owner;
            properties[FieldName.Invoices.Category] = entity.Category;
            properties[FieldName.Invoices.Status] = entity.Status;

            return properties;
        } catch (error)
        { }
    }

    public async addItem(entity: PdfViewerEntity): Promise<number> 
    {
        try
        {
            const iar: IItemAddResult = await this._sp.web.lists.getByTitle(Lists.DavidList.Name).items.add({
                Title: entity.Title, // testTitle, //
                //Owner: entity.Owner, // testOwner, //
                Category: entity.Category, // testCategory, //
                Status: entity.Status, // testStatus, //
            });

            debugger;
            return iar.data.Id;
        } catch (error)
        {
            console.error(error);
            debugger;
        }
    }
    public async updateItem(entity: PdfViewerEntity): Promise<boolean> 
    {
        throw new Error("Method not implemented.");
    }

}