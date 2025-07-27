function onOpen() {
  SlidesApp.getUi()
    .createMenu('Download')
    .addItem('All Slides as PNGs', 'saveSlidesAsPngsToDrive')
    .addToUi();
}

/**
 * Saves all slides of the active presentation as PNG images to Google Drive.
 */
function saveSlidesAsPngsToDrive() {
  const presentation = SlidesApp.getActivePresentation();
  const presentationId = presentation.getId();
  const slides = presentation.getSlides();

  // Create a folder in Google Drive to store the images
  const folderName = presentation.getName() + ' - Slide Images';
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Correctly define the optional arguments for the API call
  const thumbnailProperties = {
    'thumbnailProperties.mimeType': 'PNG',
    'thumbnailProperties.thumbnailSize': 'LARGE'
  };

  for (let i = 0; i < slides.length; i++) {
    const pageObjectId = slides[i].getObjectId();

    try {
      // Use the Slides API to get the thumbnail URL.
      // This is the JavaScript equivalent of the API method.
      const thumbnail = Slides.Presentations.Pages.getThumbnail(
        presentationId,
        pageObjectId,
        thumbnailProperties
      );
      
      const thumbnailUrl = thumbnail.contentUrl;

      // Fetch the image from the URL and save it to Drive
      const response = UrlFetchApp.fetch(thumbnailUrl);
      const imageBlob = response.getBlob();
      const fileName = `Slide ${i + 1}.png`;
      
      folder.createFile(imageBlob).setName(fileName);
      Logger.log(`Successfully saved ${fileName}`);

    } catch (e) {
      Logger.log(`Error processing slide ${i + 1}: ${e.toString()}`);
    }
  }

  SlidesApp.getUi().alert('Download Complete', 'All slides have been saved as PNG files to your Google Drive in a folder named "' + folderName + '".', SlidesApp.getUi().ButtonSet.OK);
}
