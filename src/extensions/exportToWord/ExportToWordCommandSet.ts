import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
  //RowAccessor,
  //ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Convert2Doc } from './Convert2Doc';
//import {getSP } from './pnpjsConfig';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/profiles";
//import { Dialog } from '@microsoft/sp-dialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
 export interface IExport2WordCommandSetProperties {
  listItems:[
    {
        "ID":"",
        "Kam":"";
    }
];
ID:string;

    
    
}

interface IDictionatry{
  [key: string]: any;
}

const LOG_SOURCE: string = 'Export2WordCommandSet';

export default class Export2WordCommandSet extends BaseListViewCommandSet<IExport2WordCommandSetProperties> {

  sp:any;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized Export2WordCommandSet');
    this.sp = spfi().using(SPFx(this.context));
    
    return Promise.resolve();
  }


  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const export2WordCommand: Command = this.tryGetCommand('Export2Word');
    //var listUrl = this.context.pageContext.list.title;
    if (export2WordCommand) {
      // This command should be hidden if selected any rows.
      // export2WordCommand.visible = !(event.selectedRows.length > 0);
      export2WordCommand.visible = (event.selectedRows.length===1)// && listUrl== "Denník dispečera");
    }
  }

  
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'Export2Word':
        const cnvrt2docx: Convert2Doc = new Convert2Doc(this.context.spHttpClient as any, this.context.pageContext.web.absoluteUrl, LOG_SOURCE, this.context.pageContext.list.title);
        event.selectedRows.length === 0 ? cnvrt2docx.createDocument() : this.createDocumentSelectedItems(event, cnvrt2docx);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

/**
 * Creates the documents for the selected items only
 * @param event 
 * @param cnvrt2docx 
 */

  public returnID():string{
    return this.properties.ID.toString();
  }
  private  async getUserProperties():Promise<string>{
    //const pageUrl = "https://pozfond.sharepoint.com/sites/poolcars";
    let userManager:string = "";
    let managerFullName:string = "";

    // rest api sharepoint user properties
    /*$.ajax({
        
        url: pageUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        
        method: "GET",

        headers: { "Accept": "application/json; odata=verbose" },

        success: function (data) {

            //var userProfilePropertyValue = data.d.UserProfileProperties.results.find(KeyValuePair => KeyValuePair.Key === userProfilePropertyName).Value;
            userProperties= data.d["ExtendedManagers"].results[0].split("|");
            console.log(userProperties[2]);
        
            
},

        error: function (error) {

            console.log("Error in retriving the user profile property:");

            console.log(error);

        }

    });*/
        // user properties by jsom
      
      
       /*await this.sp.profiles.myProperties.get().then(result=>{
        var userProperties = result.UserProfileProperties;
        var userPropertyValues = {};
        console.log("Manazer");
        console.log(userManager);
        console.log(userProperties[14]["Value"]);
        console.log(userPropertyValues);
        userManager += userProperties[14]["Value"];
        
        //console.log(userProperties);
        userProperties.forEach(property=> {  
            userPropertyValues[property.Key] = property.Value;  
            
        });
       });
       if(userManager!="")
       {
       await pnp.sp.profiles.getPropertiesFor(userManager)
       
                .then(result=>{
                    
                    managerFullName += result.UserProfileProperties[4]["Value"] + " " + result.UserProfileProperties[6]["Value"];
                });
        }*/

        
        const profile = await this.sp.profiles.myProperties();
        console.log("Vypisujem profile");
        console.log(profile.UserProfileProperties[14].Value);
        userManager = profile.UserProfileProperties[14].Value;

        if(userManager!=="")
        {
          await this.sp.profiles.getPropertiesFor(userManager).then((profile: any)=>{
              managerFullName+=profile.UserProfileProperties[4].Value;
              managerFullName+=" " +profile.UserProfileProperties[6].Value+",";
              
          });

        }
        
        console.log(managerFullName);
        return managerFullName;
       


    }

    
    
        

  
   
  private dateConvert(dateString:string):string
  { 
    //convert SK datumu na ENG. Pri svk datume 30.6. to zobralo ako 6.30 - invalid date
    let myArray = dateString.split(". ");
    const dateArray = [myArray[0],myArray[1]];
    const year = myArray[2].split(" ")[0];
    const timeArray = myArray[2].split(" ")[1].split(":");
    const myDate = new Date(Number(year),(Number(dateArray[1])-1),Number(dateArray[0]), Number(timeArray[0]), Number(timeArray[1]));
    
    //console.log(myDate.toLocaleString("en-US"));
    return myDate.toLocaleString();
  }
  private async createDocumentSelectedItems(event: IListViewCommandSetExecuteEventParameters, cnvrt2docx: Convert2Doc):Promise<void> {

    let html: string = '<table>';
    //let index: number = 0;
    //const values:any =[];
    let posadka:string = "";
    let spat:string="";
    let menoVodica:string="";
    let spz:string="";
    let hodiny:number=0;
    let dni:string="";
    //const valuesByName:any = [];
    //let fieldValueByName:any;
    const dict:IDictionatry={};
    //var url:string;
    let nadriadeny:string="";
    let zvysok:number=0;
    let cisloZiadanky:string="";

    /*var selectedStr = selected.map(function(item){ // loop all Objects
        return item.id; */

    
    
    
    event.selectedRows.forEach(i => {

      //html += `<tr style="height:30px"></tr>`;
      //let isAlternate: boolean = index % 2 == 0;
      
      i.fields.forEach(k => {
        
        //let value: string = '';
        //let fieldValue: any = i.getValue(k);

        //values.push(i.getValue(k));
        dict[k.internalName]=i.getValue(k);
        console.log(i.getValue(k) + ": " + k.internalName);
        
        /*switch (k.fieldType) {
            case "User":
            case "Person or Group":
              value = fieldValue && fieldValue.length > 0 ? fieldValue[0].title : '';
              break;
            case "Lookup":
              value = fieldValue && fieldValue.length > 0 ? fieldValue[0].lookupValue : '';
              break;
            case "TaxonomyFieldType":
              value = i.getValue(k).Label;
              break;
            case "URL":
              value = `<a href="${i.getValue(k)}" style="cursor:pointer;">${i.getValue(k)}</a>`;
              break;
            case "DateTime":
              value = new Date(i.getValue(k)).toLocaleString();
              //value = new Date(i.getValue(k)).toLocaleString()=="Invalid Date" ? fieldValue :"";
              //value = i.getValue(k);
              break;
            default:
              value = i.getValue(k);
          }*/
        
        
      });
      
      
    });
    console.log(dict);
    
    if(dict.acColPosadka.length > 0) {
      for(const i in dict.acColPosadka) {
      
        //posadka+=dict["acColPosadka"][i]["title"]+", ";
        
        posadka += dict.acColPosadka[i].title+", ";
        
      

  }}
    if(dict.acColSpiatocnaCesta==="Áno") {
        spat="a späť";}

    if(dict.acColVodic.length>0){
      for(const i in dict.acColVodic)  {
          menoVodica+=dict.acColVodic[i].title+" "; 
    }}
    
    if(dict.acColLookupVozidlo.length>0){
      for(const i in dict.acColLookupVozidlo)  {
          spz=dict.acColLookupVozidlo[i].lookupValue+" ";
          console.log("spz");
          console.log(spz);
    }}
   
    const d1 = new Date(this.dateConvert(dict.acColDatumCasOd));
    const d2 = new Date(this.dateConvert(dict.acColDatumCasDo));
    
    //prepocet dni
    dni+=Math.floor((Number(d2)-Number(d1))/86400000);

    //prepocet hodin, ak je recionalne cislo , zaokruhli ho
    zvysok += ((((Number(d2)-Number(d1))/1000)%86400)/3600)%1;
    if(zvysok===0)
    {
        hodiny += (((Number(d2)-Number(d1))/1000)%86400)/3600;
    }
    else{
        hodiny += Number(((((Number(d2)-Number(d1))/1000)%86400)/3600).toFixed(2));
    }
    
    this.properties.ID=dict.ID.toString();
    cisloZiadanky += `${new Date().getFullYear()}/${dict.ID}`;
    
    if(Number(dni)<1){dni="";}

    await this.getUserProperties().then((properties)=>{
        nadriadeny=properties;
    });
    console.log("nadriadeny" + nadriadeny);
    console.log(dni,hodiny);
    html+= `<table style="border-collapse:collapse;border:none;">
    <tbody>
        <tr>
            <td colspan="2" rowspan="4" style="width: 145.25pt;border-width: 1.5pt 1.5pt 1pt;border-style: solid;border-color: windowtext;border-image: initial;padding: 0in 3.5pt;height: 17.1pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:2.0pt;'><span style="font-size:11px;color:#C00000;">Organiz&aacute;cia (pečiatka)</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><span style="font-size:11px;color:#C00000;">&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><span style="font-size:11px;color:#C00000;">&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><span style="font-size:11px;color:#C00000;">&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><span style="font-size:11px;color:#C00000;">&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><span style="font-size:11px;color:#C00000;">&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:3.0pt;'><span style="font-size:11px;color:#C00000;">Žiadateľ &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><strong><span style="font-size:13px;">&nbsp;</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style='font-size:15px;font-family:"Calibri",sans-serif;color:black;'>&nbsp; &nbsp; ${dict.acColZiadatelOJ} &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span></strong></p>
            </td>
            <td colspan="2" rowspan="3" style="width: 134.7pt;border-top: 1.5pt solid windowtext;border-right: 1.5pt solid windowtext;border-bottom: 1.5pt solid windowtext;border-image: initial;border-left: none;padding: 0in 3.5pt;height: 17.1pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;text-align:center;'><strong><span style="font-size:19px;color:#C00000;">ŽIADANKA</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;text-align:center;'><strong><span style="font-size:19px;color:#C00000;">na prepravu</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="color:#C00000;">&nbsp;</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">os&ocirc;b*</span></strong><span style="font-size:11px;color:#C00000;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <s>n&aacute;kladu*</s>)</span></p>
            </td>
            <td rowspan="2" style="width: 148.8pt;border-top: 1.5pt solid windowtext;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 17.1pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:2.0pt;'><span style="font-size:11px;color:#C00000;">Č&iacute;slo objedn&aacute;vky žiadateľa</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:2.0pt;'><strong><em><span style="color:#C00000;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span></em></strong></p>
            </td>
            <td style="height:17.1pt;border:none;"><br></td>
        </tr>
        <tr>
            <td style="height:14.2pt;border:none;"><br></td>
        </tr>
        <tr>
            <td rowspan="2" style="width: 148.8pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 16.85pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:2.0pt;'><span style="font-size:11px;color:#C00000;">Č&iacute;slo objedn&aacute;vky &uacute;tvaru</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">dopravy</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:18px;font-family:"Times New Roman",serif;'><strong>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; ${cisloZiadanky}</strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;">&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span></p>
            </td>
            <td style="height:16.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="2" style="width: 134.7pt;border-top: none;border-right: none;border-left: none;border-image: initial;border-bottom: 1pt solid windowtext;padding: 0in 3.5pt;height: 0.2in;vertical-align: bottom;">
                <h2 style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:"Times New Roman",serif;'><span style="font-size:13px;color:#C00000;">&nbsp;</span></h2>
            </td>
            <td style="height:.2in;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 23.7pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Men&aacute; cestuj&uacute;cich*) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><span style="font-size:12px;">${posadka}<span style="color:#C00000;">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span></span></p>
            </td>
            <td style="height:23.7pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Druh, hmotnosť a rozmer n&aacute;kladu*) &nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><strong><em><span style="font-size:15px;color:#C00000;">&nbsp;</span></em></strong></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:135.0pt;'><strong><span style="font-size:13px;color:#C00000;">&nbsp;</span></strong></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 18.65pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Deň, hodina a miesto pristavenia*) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><span style="font-size:19px;color:black;">${d1.getDate()}.${d1.getMonth()+1}.&nbsp;-&nbsp;${d2.getDate()}.${d2.getMonth()+1}.${d2.getFullYear()}&nbsp;</span></p>
            </td>
            <td style="height:18.65pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Odkiaľ &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><span style="font-size:15px;color:black;">${dict.acColOdkial}-${dict.acColKam} &nbsp;${spat} &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Vodič sa hl&aacute;si u&nbsp;</span></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="5" style="width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Vozidlo je požadovan&eacute; na &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style="font-size:13px;color:black;">${hodiny}</span> &nbsp; &nbsp;hod&iacute;n</span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<strong><span style="font-size:11px;">&nbsp; &nbsp; &nbsp;</span></strong><span style="font-size:11px;color:#C00000;"><span style="font-size:13px;color:black;">${dni}</span>&nbsp; &nbsp; dni &nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-left:.25in;'><span style="font-size:11px;color:#C00000;">&nbsp;</span></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="3" style="width: 185.35pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 20.4pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">&Uacute;čel jazdy &nbsp;</span><span style="font-size:13px;color:black;">${dict.Title},&nbsp;</span></p>
            </td>
            <td colspan="2" rowspan="2" style="width: 243.4pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 20.4pt;vertical-align: top;">
                <h1 style='margin:0in;margin-bottom:.0001pt;text-align:center;font-size:21px;font-family:"Times New Roman",serif;font-weight:normal;'><strong><span style="font-size:15px;color:#C00000;border:solid windowtext 1.0pt;padding:0in;background:white;">PR&Iacute;KAZ NA JAZDU</span></strong><span style="font-size:15px;color:#C00000;border:solid windowtext 1.0pt;padding:0in;background:white;">&nbsp; &nbsp;</span><span style="font-size:15px;color:#C00000;background:  white;">&nbsp;&nbsp;</span></h1>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:4.0pt;'><span style="font-size:11px;color:#C00000;">Meno vodiča &nbsp; &nbsp;&nbsp;</span><span style="color:black;">${menoVodica}</span></p>
            </td>
            <td style="height:20.4pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="3" style="width: 185.35pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 27.8pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Vy&uacute;čtuje na vrub &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><strong><span style="font-size:15px;">Ekonomick&yacute; odbor</span></strong></p>
            </td>
            <td style="height:27.8pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="3" rowspan="2" style="width: 185.35pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;margin-top:2.0pt;'><span style="font-size:11px;color:#C00000;">Pozn&aacute;mka žiadateľa :&nbsp;</span><span style="font-size:13px;color:black;">${dict.acColPoznamka},&nbsp;</span></p></p>
            </td>
            <td colspan="2" style="width: 243.4pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">Druh vozidla &nbsp; &nbsp;&nbsp;</span><span style="font-size:11px;color:black;">${dict.Vozidlo_x003a_Druh_x0020_vozidla}</span></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td colspan="2" style="width: 243.4pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">&Scaron;PZ &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><strong><span style="font-size:13px;">${spz}</span></strong></p>
            </td>
            <td style="height:11.85pt;border:none;"><br></td>
        </tr>
        <tr>
            <td style="width: 92.45pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1.5pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 47.25pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">D&aacute;tum a podpis&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">žiadateľa &nbsp;&nbsp;</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">${nadriadeny}</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">${d1.getDate()}.${d1.getMonth()+1}.${d1.getFullYear()}</span></strong></p>
            </td>
            <td colspan="2" style="width: 92.9pt;border-top: none;border-left: none;border-bottom: 1.5pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 47.25pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">D&aacute;tum a&nbsp;podpis</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">schvaľuj&uacute;ceho</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">Ing. Puchelov&aacute;</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">${d1.getDate()}.${d1.getMonth()+1}.${d1.getFullYear()}</span></strong></p>
            </td>
            <td colspan="2" style="width: 243.4pt;border-top: none;border-left: none;border-bottom: 1.5pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 47.25pt;vertical-align: top;">
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">D&aacute;tum a podpis osoby zodpovednej</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><span style="font-size:11px;color:#C00000;">za autoprev&aacute;dzku</span></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">Peter &Scaron;tetina</span></strong></p>
                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:"Times New Roman",serif;'><strong><span style="font-size:11px;">${d1.getDate()}.${d1.getMonth()+1}.${d1.getFullYear()}</span></strong></p>
            </td>
            <td style="height:47.25pt;border:none;"><br></td>
        </tr>
    </tbody>
</table>`;
    
    
    
    console.log("cisloZiadanky - return");
    console.log( this.returnID());
    //cnvrt2docx.cisloZiadanky=cisloZiadanky;
    cnvrt2docx.CisloZiadanky=cisloZiadanky;
    await cnvrt2docx.generateDocument(html);

  }
}