export class Person
{
    public Id: string = '';

    public Title: string = '';

    public Email: string = '';
}

export class PdfViewerEntity
{
 public Id: number = -1;
 
 public Title: string = '';



 public Owner: Person = new Person();

 public Status: string = ''; //Choice Field

 public Category: string = ''; //Choice Field

 //public Demo Image: string = '';
}

