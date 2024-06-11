using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Azure.Storage.Sas;
using PowerPointApi;

// Alias the conflicting types to differentiate them
using DShapeProperties = DocumentFormat.OpenXml.Drawing.ShapeProperties;
using PShapeProperties = DocumentFormat.OpenXml.Presentation.ShapeProperties;

namespace PowerPointApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PowerPointController : ControllerBase
    {
        private readonly AzureSettings _azureSettings;

        public PowerPointController(IOptions<AzureSettings> azureSettings)
        {
            _azureSettings = azureSettings.Value;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadPowerPoint(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("The file field is required.");
            }

            var tempFilePath = System.IO.Path.GetTempFileName();

            using (var stream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
            {
                await file.CopyToAsync(stream);
            }

            // Perform PowerPoint manipulations
            var pptManipulation = new PowerPointManipulation();
            pptManipulation.MakeAllBackgroundsTransparent(tempFilePath);

            // Upload to Azure Blob Storage
            var blobUri = await UploadToAzureBlob(tempFilePath, file.FileName);

            // Return the file link
            return Ok(new { FileLink = blobUri });
        }

        private async Task<string> UploadToAzureBlob(string filePath, string fileName)
        {
            var blobServiceClient = new BlobServiceClient(_azureSettings.ConnectionString);
            var blobContainerClient = blobServiceClient.GetBlobContainerClient(_azureSettings.ContainerName);
            
            await blobContainerClient.CreateIfNotExistsAsync();

            var blobClient = blobContainerClient.GetBlobClient(fileName);

            await blobClient.UploadAsync(filePath, new BlobHttpHeaders { ContentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation" });

            // Generate a SAS token for the blob
            var sasToken = blobClient.GenerateSasUri(BlobSasPermissions.Read, DateTimeOffset.UtcNow.AddHours(1));
            
            return sasToken.ToString();
        }
    }

    public class PowerPointManipulation
    {
        public void MakeAllBackgroundsTransparent(string filePath)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, true))
            {
                // Process master slides and their layouts
                foreach (var slideMasterPart in presentationDocument.PresentationPart.SlideMasterParts)
                {
                    ClearBackgroundProperties(slideMasterPart.SlideMaster.CommonSlideData?.Background);
                    RemoveBackgroundElements(slideMasterPart.SlideMaster.CommonSlideData?.ShapeTree);

                    // Process each layout associated with this master
                    foreach (var layoutPart in slideMasterPart.SlideLayoutParts)
                    {
                        ClearBackgroundProperties(layoutPart.SlideLayout.CommonSlideData?.Background);
                        RemoveBackgroundElements(layoutPart.SlideLayout.CommonSlideData?.ShapeTree);
                    }
                }

                // Process each individual slide
                foreach (var slidePart in presentationDocument.PresentationPart.SlideParts)
                {
                    ClearBackgroundProperties(slidePart.Slide.CommonSlideData?.Background);
                    RemoveBackgroundElements(slidePart.Slide.CommonSlideData?.ShapeTree);
                }

                presentationDocument.PresentationPart.Presentation.Save(); // Save the modified presentation
            }
        }

        private void ClearBackgroundProperties(Background background)
        {
            if (background != null)
            {
                background.RemoveAllChildren<BackgroundProperties>();
                var backgroundProperties = new BackgroundProperties();
                var solidFill = new SolidFill(new Alpha { Val = 0 }); // Set transparency to 100%
                backgroundProperties.Append(solidFill);
                background.Append(backgroundProperties);
            }
        }

        private void RemoveBackgroundElements(ShapeTree shapeTree)
        {
            if (shapeTree != null)
            {
                foreach (var element in shapeTree.Descendants<DocumentFormat.OpenXml.OpenXmlElement>())
                {
                    if (element is DocumentFormat.OpenXml.Presentation.Shape shapeElement && shapeElement.ShapeProperties != null)
                    {
                        // Remove any fill properties from shapes
                        shapeElement.ShapeProperties.RemoveAllChildren<Fill>();
                        shapeElement.ShapeProperties.Append(new NoFill());
                    }

                    if (element is DocumentFormat.OpenXml.Presentation.Picture picture)
                    {
                        // Remove pictures that may be used as backgrounds
                        picture.Remove();
                    }

                    // Handle grouped shapes
                    if (element is DocumentFormat.OpenXml.Presentation.GroupShape groupShape)
                    {
                        foreach (var childShape in groupShape.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
                        {
                            if (childShape.ShapeProperties != null)
                            {
                                childShape.ShapeProperties.RemoveAllChildren<Fill>();
                                childShape.ShapeProperties.Append(new NoFill());
                            }
                        }
                    }
                }
            }
        }
    }
}
