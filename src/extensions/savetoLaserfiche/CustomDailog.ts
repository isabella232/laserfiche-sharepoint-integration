import { BaseDialog,IDialogConfiguration} from '@microsoft/sp-dialog';
import styles from './SendToLaserFiche.module.scss';
//import {} from './../../../lib/logo/laserfiche-logo.png';

export default class CustomDailog extends BaseDialog {  
    
    public render(): void {  
        let html = "";  
        html+=  `<div  class="${ styles["maindialog"] }">`; 
        html += `<div id="overlay" class="${ styles.overlay }"></div>`;
        //html+=  `<div>`;
        html+=  `<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAALQAAAC0CAMAAAAKE/YAAAAAUVBMVEXSXyj////HYzL/+/T/+Or/9d+yaUa9ZT2yaUj/9OG7Zj3SXybRYCj/+/b///3LYS/OYCvEZDS2aEL/89jAZTnMYS3/8dO7Zzusa02+ZTn/78wyF0DsAAABnUlEQVR4nO3ci26CMABGYQcoLRS5OTf2/g86R+KSLYUm2vxcPB8RTYzxkADRajkcAAAAAAAAAADYgbJcusCvqdtLnhfeJR/a96X7vOriarNJ/cUtHeiTnI7p26TsY+XRZ190sXSfVyA6X7rP6xZdzeweREeTGDt3IBIdTeCUR3Q0wQOxLNf3CWSr0ZvcPYiWIFqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVYhWIVqFaBWiVV4zeok/379m9BL2HO1Ckymlky0jRQc3Kqoou4f6YHzdaLX56PRzak757/JjfDS0dbOK6HM6Paf8P3st6lVE/9mAwPOpNcnqokOIJppoookmmmiiiSaaaKKJ3k30OfTFdU3RXZ+lT6qq6rbO+k4VXQ9fvT2OrH30Zo+3u/5rUI17NO3QmdPImIduxoyrUze0khEm5w6uqZNIRKNi91Hl5661dH+tdow6wts5J//BaJPRwH6IT1NxbDJ6vVc+nrXJaAAAAADALn0DBosqnCStFi4AAAAASUVORK5CYII=" width="42" height="42">`;
        //html+=  `</div>`;
        html += `<img style="margin-left:28%;" src="/_layouts/15/images/progress.gif" id="imgid">`;
        html+=  `<div>`;
        html+=  `<p class="${ styles.text }" id="it">Saving your document to Laserfiche</p>`;
        html+=  `</div>`;
        html+=  `<div id="divid" class="${ styles.button }">`;
        html+=  `<button id="divid1" class="${ styles.button1 }">Close</button>`;
        html+=  `<button id="divid13" class="${ styles.button2 }" title="Click here to view the file in Laserfiche">View File</button>`;
        html+=  `</div>`;
        //html+=  `<div id="divid12" class="${ styles.button }">`;
        //html+=  `<button id="divid13" class="${ styles.button1 }" title="Click here to view File in Laserfiche">View File</button>`;
        //html+=  `</div>`;
        html+=  `</div>`;
        /* html+= ` <div className="card">
        <div className="lf-component-container">
          <lf-field-container collapsible="true" startCollapsed="true" ref={this.fieldContainer}>
            </lf-field-container>
          </div>
        </div>` */
        this.domElement.innerHTML = html;       
    }  
   public  getConfig(): IDialogConfiguration {
    return {
    isBlocking: false
    };
}
    protected onAfterClose(): void {  
      super.onAfterClose();       
    }     
  } 