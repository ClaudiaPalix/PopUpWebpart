import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './PopUpApplicationCustomizer.module.scss';

import * as strings from 'PopUpApplicationCustomizerStrings';
import { concatStyleSets } from 'office-ui-fabric-react';

const LOG_SOURCE: string = 'PopUpApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPopUpApplicationCustomizerProperties {
  userEmail: string;
  company: string;
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PopUpApplicationCustomizer
  extends BaseApplicationCustomizer<IPopUpApplicationCustomizerProperties> {

    private slideIndex: number = 1;
    private userEmail: string = "";
    private Slides: any = "";
    private renderSlides = false;

    private async userDetails(): Promise<void> {
      // Ensure that you have access to the SPHttpClient
      const spHttpClient: SPHttpClient = this.context.spHttpClient;
    
      // Use try-catch to handle errors
      try {
        // Get the current user's information
        const response: SPHttpClientResponse = await spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1);
        const userProperties: any = await response.json();
    
        // console.log("User Details:", userProperties);
    
        // Access the userPrincipalName from userProperties
        const userPrincipalNameProperty = userProperties.UserProfileProperties.find((property: any) => property.Key === 'SPS-UserPrincipalName');
    
        if (userPrincipalNameProperty) {
          this.userEmail = userPrincipalNameProperty.Value.toLowerCase();
          console.log('User Email using User Principal Name:', this.userEmail);
          // Now you can use this.userEmail as needed
        } else {
          console.error('User Principal Name not found in user properties');
        }
      } catch (error) {
        console.error('Error fetching user properties:', error);
      }
    } 


  public async onInit(): Promise<void> {
    // You can trigger your popup logic here
    const currentUrl = window.location.href;
    // console.log("Current Url:", currentUrl);
    // console.log("Page Context:", this.context.pageContext.web.absoluteUrl);
    // console.log("currentUrl === this.context.pageContext.web.absoluteUrl", currentUrl === this.context.pageContext.web.absoluteUrl)
    if ((currentUrl === this.context.pageContext.web.absoluteUrl) || (currentUrl === "https://rpgnet.sharepoint.com/sites/OneRPG/SitePages/HomeTeams.aspx")) {
    await this.userDetails();
    this.Slides = await this.getSlides();
    if (!this.Slides.value || this.Slides.value.length === 0) {
      console.log("No slides available. Exiting program.");
      return Promise.resolve();
    }
    await this.getAttachments();
    console.log("Slides with attachments:", this.Slides);
    if (this.renderSlides) {
      await this.showPopup();
    }
    }
    return Promise.resolve();
  }
  
  private async getSlides(): Promise<any> {
      let userEmail = this.context.pageContext.user.email.toLowerCase();
      // console.log("User Email:", userEmail);
      if(!userEmail){
        let extracted = this.properties.userEmail.split("#ext#")[0];
        let lastUnderscoreIndex = extracted.lastIndexOf("_");
        if (lastUnderscoreIndex !== -1) {
          extracted = extracted.substring(0, lastUnderscoreIndex) + "@" + extracted.substring(lastUnderscoreIndex + 1);
        }
        userEmail = extracted;
        // console.log(extracted);
      }
  
      let match = userEmail.match(/@([\w.-]+)\.(com|in)/);
      let company = match ? match[1] : null;
      console.log("User's Company",company); 
  
      const today = new Date();
      const tomorrow = new Date();
      tomorrow.setDate(today.getDate() - 1);
      // const dateOnly = today.toLocaleDateString();
      const isoStringToday = today.toISOString();
      const dateAndTimeToday = isoStringToday.substring(0, isoStringToday.length - 5) + "Z";
      const isoStringTomorrow = tomorrow.toISOString();
      const dateAndTimeTomorrow = isoStringTomorrow.substring(0, isoStringTomorrow.length - 5) + "Z";
      // console.log(dateAndTimeToday);
      // console.log(dateAndTimeTomorrow);
      // console.log(today); 
      // console.log(dateOnly); // Outputs: "mm/dd/yyyy" in US locale
      let apiUrl: string =``;
  
      if(this.userEmail.includes(".admin@")){
      //if admin then show all popups no audience targeting
      apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PopUp_List')/items?filter=StartDate%20le%20%27${dateAndTimeToday}%27%20and%20EndDate%20ge%20%27${dateAndTimeTomorrow}%27&$orderby=Modified desc`; 
      }else{
        apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PopUp_List')/items?$filter=(StartDate%20le%20%27${dateAndTimeToday}%27%20and%20EndDate%20ge%20%27${dateAndTimeTomorrow}%27)%20and%20(AudienceTarget%20eq%20%27${company}%27%20or%20AudienceTarget%20eq%20%27all%20companies%27)&$orderby=Modified%20desc`;  
      }
  
      // console.log("Api Url:",apiUrl);
      const response = await fetch(apiUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      });
    
      if (!response.ok) {
        console.error('Error fetching slides:', response.statusText);
        return;
      }
    
      const data = await response.json();
      // console.log("Slides Api response:", data);
    
      return data; // Add this line
  }

  private async getAttachments(): Promise<any> {

      if (!this.Slides || !this.Slides.value) {
        console.error('this.Slides or this.Slides.value is undefined');
        return;
      }
      // else{
        // console.log("this.Slides is defined")
      // }
      const attachmentPromises = this.Slides.value.map(async (slide: any) => {
          const attachmentUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('PopUp_List')/items(${slide.Id})/AttachmentFiles`;
          const attachmentResponse = await fetch(attachmentUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          });
      
          if (!attachmentResponse.ok) {
            console.error('Error fetching attachments:', attachmentResponse.statusText);
            return;
          }
      
          const attachmentData = await attachmentResponse.json();
          // console.log("Attachments for item " + slide.Id + ":", attachmentData);
      
          // Add the attachments to the slide
          // console.log('attachmentData.value:', attachmentData.value);
          slide.Attachment = attachmentData.value;
          // console.log('slide after setting Attachments:', slide);    
          return slide;    
        });

         // Wait for all promises to complete
  const slidesWithAttachments = await Promise.all(attachmentPromises);
  // console.log("Slides with attachments:", slidesWithAttachments);

  // Update this.Slides with the slides that have attachments
  this.Slides.value = slidesWithAttachments;

  // Set this.renderSlides to true
  this.renderSlides = true;
}

  private async _renderSlides():Promise<void>{
    const domain = "https://rpgnet.sharepoint.com";
    // console.log("Slides after all calls:", this.Slides);
    const OuterDiv: HTMLElement | null = document.querySelector(`#slideContainer`);
    const TextDiv= document.createElement('div');
    TextDiv.classList.add(styles['textDiv']);
    const Div = document.createElement('div');
    Div.classList.add(styles['slideshow-container']);

    let extension = "";
    if (!OuterDiv || !Div) {
      console.error("Element in PopUp not found");
      return;
    }else{
      console.log("Outer Div:",OuterDiv);
    }
    const data = this.Slides; 
    const length = data.value.length;
    let slideNumber = 0;
    if(data.value.length > 0){
      data.value.forEach((slide: any) => {
        slideNumber++;

        const slideDiv = document.createElement('div');
        slideDiv.classList.add(styles.fade);
        slideDiv.classList.add(styles.popupSlides);
        // console.log("Slide:",slide);
          // console.log("Slide Attachements:",slide.Attachment);
          const fullImage = `${domain}${slide.Attachment[0].ServerRelativeUrl}`;
        console.log("Full Image:",fullImage);
        slideDiv.innerHTML = `
                      <div class="${styles.numbertext}">${slideNumber} / ${length}</div>
                    `;
        // Get the file extension
        let tempExtension = fullImage.split('.').pop();

        // Check if tempExtension is not undefined
        if (tempExtension) {
          extension = tempExtension;
        } else {
          // Default to image if no extension found
          extension = 'jpg';
        }

        // Convert extension to lower case
        extension = extension.toLowerCase();

        let content;
        if (['jpg', 'jpeg', 'png', 'gif'].includes(extension)) {
          // If the file is an image, use an img tag
          content = `<img src="${fullImage}" alt="Slide image" style="width:100%"/>`;
        slideDiv.innerHTML = content;
        } else if (extension === 'mp4') {
          // If the file is a video, use a video tag
          // Create the video element
          let videoElement = document.createElement('video');
          videoElement.controls = true;
          videoElement.style.width = '100%';

          let sourceElement = document.createElement('source');
          sourceElement.src = fullImage;
          sourceElement.type = 'video/mp4';
          sourceElement.style.width = '100%';

          videoElement.appendChild(sourceElement);
          slideDiv.appendChild(videoElement);

          // Create an intersection observer
          let observer = new IntersectionObserver((entries, observer) => {
            entries.forEach(entry => {
              // If the video is in the viewport, play it
              if (entry.isIntersecting) {
                (entry.target as HTMLVideoElement).play();
              } else {
                // If the video is not in the viewport, pause it
                (entry.target as HTMLVideoElement).pause();
              }
            });
          }, { threshold: 0.5 });  // The callback will run when 50% of the video is visible

          // Start observing the video
          observer.observe(videoElement);
        } else {
          // If the file type is not supported, display a message
          content = `<p>File type not supported</p>`;
        slideDiv.innerHTML = content;
        }

        const bottomGradient = document.createElement('div');
        bottomGradient.classList.add(styles.bottomGradient);
        if (extension != 'mp4') {
          bottomGradient.onclick = () => {
            window.open(slide.LinkUrl.Url, '_blank');
            };        
          slideDiv.appendChild(bottomGradient);
        }else{
        bottomGradient.style.display = "none";
        }
        
        Div.appendChild(slideDiv);
        if(slide.Description){
          const slideText = document.createElement('div');
          slideText.classList.add(styles.text);
          slideText.innerHTML = slide.Description;
          // console.log("Slide Link:",slide.LinkUrl.Url); 
          slideText.onclick = () => {
            window.open(slide.LinkUrl.Url, '_blank');
            };        
          slideDiv.appendChild(bottomGradient);
          TextDiv.appendChild(slideText);
        }
      });
    }else{
      console.log("No slides found");
    }
    

    const prev = document.createElement('a');
    prev.classList.add(styles.prev);
    prev.innerHTML = '&#10094;';
    // prev.onclick = () => this.plusSlides(-1);
    Div.appendChild(prev);
    const next = document.createElement('a'); 
    next.classList.add(styles.next);
    next.innerHTML = '&#10095;';
    // next.onclick = () => this.plusSlides(1);
    Div.appendChild(next);
    
    const dotsDiv = document.createElement('div');
    dotsDiv.classList.add(styles.dotButton);
    
    for(let index = 0; index < length; index++) {
      // console.log("Index:",index+1);
      let dotElement = document.createElement('span');
      dotElement.setAttribute("data-slide-index", `${index+1}`);
      dotElement.classList.add(styles.dot);
      // dotElement.onclick = () => this.currentSlide(index+1);
      dotsDiv.appendChild(dotElement); 
    }

    Div.appendChild(dotsDiv);
    OuterDiv.appendChild(Div);
    OuterDiv.appendChild(TextDiv);
  }

  private async showPopup(): Promise<void> {
    const modal = document.createElement('div');
    modal.id = 'popupModal';
    modal.classList.add(styles['modal']);
    modal.innerHTML = `
    
  <div class="${styles['modal-content']}">

  <span id="close" class="${styles.close}">&times;</span>

  <div id="modalBody" class="${styles['modal-body']}">
    
      <div id="slideContainer">      
  
      </div>

  </div>
</div>
    `;

    document.body.appendChild(modal);

    modal.style.display = 'block';

    // (x), close the modal
    const closeBtn = modal.querySelector('#close');
    // console.log("Close Button:", closeBtn);
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        modal.style.display = 'none';
      });
    }

    // When the user clicks anywhere outside of the modal, close it
    window.addEventListener('click', (event) => {
      if (event.target === modal) {
        modal.style.display = 'none';
      }
    });

  this.setupEventHandlers();
  }

  private async setupEventHandlers(): Promise<void> {
    
    await this._renderSlides();
    setTimeout(() => {
      this.showSlides(this.slideIndex);
    }, 0);

    document.addEventListener('click', (event: Event) => {
      const target = event.target as HTMLElement;
      if (target.classList.contains(`${styles.prev}`)) {
        this.plusSlides(-1);
      } else if (target.classList.contains(`${styles.next}`)) {
        this.plusSlides(1);
      } else if (target.classList.contains(`${styles.dot}`)) {
        const index = parseInt(target.getAttribute('data-slide-index') || '1', 10);
        this.currentSlide(index);
      }
    });
  
    setInterval(() => {
      this.plusSlides(1);
    }, 4000);

    // console.log("End of setupEventHandlers");
  }

  private plusSlides(n: number): void {
    // console.log("Start of plusSlides");
    this.showSlides(this.slideIndex + n);
    // console.log("End of plusSlides");
  }

  private currentSlide(n: number): void {
    // console.log("Start of currentSlide");
    this.showSlides(n);
    // console.log("End of currentSlide");
  }

  private showSlides(n: number): void {
    // console.log("Start of showSlides");

    const slides = document.getElementsByClassName(`${styles.popupSlides}`) as HTMLCollectionOf<HTMLElement>;
    const text = document.getElementsByClassName(`${styles.text}`) as HTMLCollectionOf<HTMLElement>;
    const dots = document.getElementsByClassName(`${styles.dot}`) as HTMLCollectionOf<HTMLElement>;

    // console.log("Slides:", slides);
    // console.log("Dots:", dots);

    if (!slides || slides.length === 0) {
      // console.error("No slides found");
      return;
    }

    // Adjusting index calculation
    if (n > slides.length) {
      this.slideIndex = 1;
    } else if (n < 1) {
      this.slideIndex = slides.length;
    } else {
      this.slideIndex = n;
    }

    // console.log("Slide Index:", this.slideIndex);
    for (let i = 0; i < slides.length; i++) {
      slides[i].style.display = "none";
      text[i].style.display = "none";
    }

    for (let i = 0; i < dots.length; i++) {
      dots[i].classList.remove(`${styles.active}`);
    }

    slides[this.slideIndex - 1].style.display = "block";
    text[this.slideIndex - 1].style.display = "block";
    dots[this.slideIndex - 1].classList.add(`${styles.active}`);

    // console.log("End of showSlides");
}

}
