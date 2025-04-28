# DOCX Conversion Libraries for .NET and Angular/Froala Integration

## .NET Backend Libraries

### 1. Open XML SDK
- **Link**: [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK)
- **Description**: Official Microsoft library for working with Office Open XML files
- **Features**:
  - Low-level access to DOCX file structure
  - Pure .NET implementation (C#)
  - High performance and accuracy
  - Full control over document structure
- **Use case**: Perfect for extracting structured content from DOCX for MongoDB storage
- **Sample code**:
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;

public class DocxProcessor
{
    public string ConvertDocxToJson(string docxPath)
    {
        var documentStructure = new DocumentStructure();
        
        using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
        {
            Body body = doc.MainDocumentPart.Document.Body;
            
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var para = new ParagraphStructure
                {
                    Text = paragraph.InnerText,
                    Style = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value
                };
                
                // Extract formatting data
                foreach (var run in paragraph.Elements<Run>())
                {
                    if (run.RunProperties != null)
                    {
                        if (run.RunProperties.Bold != null)
                            para.Formatting.Bold = true;
                        if (run.RunProperties.Italic != null)
                            para.Formatting.Italic = true;
                        // Add other formatting properties
                    }
                }
                
                documentStructure.Paragraphs.Add(para);
            }
            
            // Process tables, images, etc.
        }
        
        return JsonConvert.SerializeObject(documentStructure);
    }
}

public class DocumentStructure
{
    public List<ParagraphStructure> Paragraphs { get; set; } = new List<ParagraphStructure>();
    public List<TableStructure> Tables { get; set; } = new List<TableStructure>();
    public List<ImageStructure> Images { get; set; } = new List<ImageStructure>();
}

public class ParagraphStructure
{
    public string Text { get; set; }
    public string Style { get; set; }
    public FormattingStructure Formatting { get; set; } = new FormattingStructure();
}

// Other structure classes...
```

### 2. DocX
- **Link**: [DocX](https://github.com/xceedsoftware/DocX)
- **Description**: Higher-level .NET library for DOCX manipulation
- **Features**:
  - Simpler API than Open XML SDK
  - Good for creating and modifying DOCX files
  - Less verbose code
- **Use case**: Good for simpler documents and when you need a higher-level API
- **Sample code**:
```csharp
using Xceed.Words.NET;
using Newtonsoft.Json;

public class DocxSimpleProcessor
{
    public string ConvertDocxToJson(string docxPath)
    {
        var documentStructure = new DocumentStructure();
        
        using (DocX document = DocX.Load(docxPath))
        {
            foreach (var paragraph in document.Paragraphs)
            {
                var para = new ParagraphStructure
                {
                    Text = paragraph.Text,
                    Style = paragraph.StyleName
                };
                
                // Extract formatting is more limited with DocX
                
                documentStructure.Paragraphs.Add(para);
            }
            
            // Process tables, images, etc.
        }
        
        return JsonConvert.SerializeObject(documentStructure);
    }
}
```

### 3. Aspose.Words for .NET
- **Link**: [Aspose.Words](https://products.aspose.com/words/net/)
- **Description**: Comprehensive commercial library for document processing
- **Features**:
  - High-fidelity document processing
  - Rich conversion options (DOCX to HTML, PDF, JSON)
  - Advanced document manipulation
  - Built-in support for MongoDB integration
- **Use case**: Enterprise solutions requiring robust document handling
- **Sample code**:
```csharp
using Aspose.Words;
using Aspose.Words.Saving;

public class AsposeDocProcessor
{
    public string ConvertDocxToJson(string docxPath)
    {
        // Load the document
        Document doc = new Document(docxPath);
        
        // Set up JSON save options
        JsonSaveOptions options = new JsonSaveOptions
        {
            ExportDocumentStructure = true,
            PrettyFormat = true
        };
        
        // Convert to JSON (memory stream used for demonstration)
        using (MemoryStream stream = new MemoryStream())
        {
            doc.Save(stream, options);
            stream.Position = 0;
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
    
    public string ConvertDocxToHtml(string docxPath)
    {
        // Load the document
        Document doc = new Document(docxPath);
        
        // Set up HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true,
            ExportFontsAsBase64 = true,
            CssStyleSheetType = CssStyleSheetType.Embedded
        };
        
        // Convert to HTML
        using (MemoryStream stream = new MemoryStream())
        {
            doc.Save(stream, options);
            stream.Position = 0;
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
```

### 4. DocumentFormat.OpenXml with MongoDB Integration
- **Description**: Combined approach using OpenXML SDK with MongoDB.Driver
- **Features**:
  - Custom mapping between DOCX structure and MongoDB documents
  - Efficient storage and retrieval
- **Sample code**:
```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MongoDB.Bson;
using MongoDB.Driver;

public class DocxMongoService
{
    private readonly IMongoCollection<BsonDocument> _documentsCollection;
    
    public DocxMongoService(IMongoDatabase database)
    {
        _documentsCollection = database.GetCollection<BsonDocument>("documents");
    }
    
    public async Task<string> StoreDocxAsStructuredData(string docxPath, string title)
    {
        var structuredDoc = new BsonDocument
        {
            { "title", title },
            { "createdDate", DateTime.UtcNow },
            { "lastModified", DateTime.UtcNow }
        };
        
        var contentDoc = new BsonDocument();
        var paragraphs = new BsonArray();
        
        using (WordprocessingDocument doc = WordprocessingDocument.Open(docxPath, false))
        {
            // Extract document properties/metadata
            if (doc.CoreFilePropertiesPart != null)
            {
                var props = doc.CoreFilePropertiesPart.GetXmlDocument();
                structuredDoc["metadata"] = new BsonDocument
                {
                    { "author", props.Descendants().FirstOrDefault(p => p.LocalName == "creator")?.InnerText },
                    { "created", props.Descendants().FirstOrDefault(p => p.LocalName == "created")?.InnerText }
                    // Add more metadata as needed
                };
            }
            
            // Process document body
            Body body = doc.MainDocumentPart.Document.Body;
            
            foreach (var paragraph in body.Elements<Paragraph>())
            {
                var paraDoc = new BsonDocument
                {
                    { "text", paragraph.InnerText },
                    { "style", paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value }
                };
                
                // Extract formatting
                var formattingDoc = new BsonDocument();
                foreach (var run in paragraph.Elements<Run>())
                {
                    if (run.RunProperties != null)
                    {
                        if (run.RunProperties.Bold != null)
                            formattingDoc["bold"] = true;
                        if (run.RunProperties.Italic != null)
                            formattingDoc["italic"] = true;
                        // Add other formatting properties
                    }
                }
                
                paraDoc["formatting"] = formattingDoc;
                paragraphs.Add(paraDoc);
            }
            
            contentDoc["paragraphs"] = paragraphs;
            
            // Add tables, images processing here
        }
        
        structuredDoc["content"] = contentDoc;
        
        // Generate HTML for Froala
        string htmlContent = await ConvertStructuredToHtml(structuredDoc);
        structuredDoc["htmlCache"] = htmlContent;
        structuredDoc["htmlCacheUpdated"] = DateTime.UtcNow;
        
        await _documentsCollection.InsertOneAsync(structuredDoc);
        return structuredDoc["_id"].AsObjectId.ToString();
    }
    
    private async Task<string> ConvertStructuredToHtml(BsonDocument structuredDoc)
    {
        // Implement conversion from structured data to HTML
        var html = new StringBuilder("<div>");
        
        var paragraphs = structuredDoc["content"]["paragraphs"].AsBsonArray;
        foreach (var para in paragraphs)
        {
            var style = para["style"].AsString;
            var text = para["text"].AsString;
            
            if (style?.StartsWith("Heading") == true)
            {
                var level = int.Parse(style.Replace("Heading", ""));
                html.AppendLine($"<h{level}>{text}</h{level}>");
            }
            else
            {
                html.AppendLine($"<p>{text}</p>");
            }
        }
        
        html.AppendLine("</div>");
        return html.ToString();
    }
}
```

## Angular/Froala Frontend Libraries

### 1. Mammoth.js
- **Link**: [Mammoth.js](https://github.com/mwilliamson/mammoth.js)
- **Description**: Pure JavaScript library for converting DOCX to HTML
- **Features**:
  - Client-side conversion
  - Good HTML output quality
  - Style mapping customization
- **Use case**: Direct DOCX to HTML conversion in browser
- **Sample Angular service**:
```typescript
import { Injectable } from '@angular/core';
import * as mammoth from 'mammoth';

@Injectable({
  providedIn: 'root'
})
export class DocxConverterService {
  
  constructor() { }
  
  async convertDocxFileToHtml(file: File): Promise<string> {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    return result.value;
  }
  
  async convertDocxFileToStructured(file: File): Promise<any> {
    // First convert to HTML
    const html = await this.convertDocxFileToHtml(file);
    
    // Then parse HTML into structured data
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    
    const structuredData = {
      paragraphs: [],
      tables: []
    };
    
    // Process paragraphs
    doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach(element => {
      const isHeading = element.tagName.toLowerCase().startsWith('h');
      const headingLevel = isHeading ? parseInt(element.tagName.substring(1)) : null;
      
      structuredData.paragraphs.push({
        text: element.textContent,
        style: isHeading ? `Heading${headingLevel}` : 'Normal',
        formatting: {
          bold: element.querySelector('strong, b') !== null,
          italic: element.querySelector('em, i') !== null,
          underline: element.querySelector('u') !== null
        }
      });
    });
    
    // Process tables
    doc.querySelectorAll('table').forEach(tableElement => {
      const table = { rows: [] };
      
      tableElement.querySelectorAll('tr').forEach(row => {
        const cells = [];
        row.querySelectorAll('td, th').forEach(cell => {
          cells.push({
            text: cell.textContent,
            isHeader: cell.tagName.toLowerCase() === 'th'
          });
        });
        
        table.rows.push({ cells });
      });
      
      structuredData.tables.push(table);
    });
    
    return structuredData;
  }
}
```

### 2. Froala Editor Integration Service
- **Description**: Custom Angular service for Froala integration with MongoDB
- **Features**:
  - Handles conversion between structured data and Froala HTML
  - Manages MongoDB API integration
- **Sample Angular service**:
```typescript
import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable } from 'rxjs';
import { map } from 'rxjs/operators';

@Injectable({
  providedIn: 'root'
})
export class FroalaDocumentService {
  private apiUrl = 'https://your-api-domain.com/api/documents';
  
  constructor(private http: HttpClient) { }
  
  // Load document from MongoDB via .NET API
  getDocument(documentId: string): Observable<any> {
    return this.http.get<any>(`${this.apiUrl}/${documentId}`)
      .pipe(
        map(response => {
          // Check if we have a cached HTML version
          if (response.htmlCache) {
            response.htmlForFroala = response.htmlCache;
          } else {
            // Convert structured content to HTML (fallback)
            response.htmlForFroala = this.convertStructuredToHtml(response.content);
          }
          return response;
        })
      );
  }
  
  // Save document from Froala to MongoDB via .NET API
  saveDocument(documentId: string, froalaHtml: string): Observable<any> {
    // Convert Froala HTML to structured content
    const structuredContent = this.convertHtmlToStructured(froalaHtml);
    
    // Send both to backend for storage
    return this.http.put<any>(`${this.apiUrl}/${documentId}`, {
      structuredContent: structuredContent,
      htmlCache: froalaHtml
    });
  }
  
  // Import a new DOCX file
  importDocx(file: File, title: string): Observable<any> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('title', title);
    
    return this.http.post<any>(`${this.apiUrl}/import`, formData);
  }
  
  // Convert structured data to HTML for Froala
  private convertStructuredToHtml(content: any): string {
    let html = '<div>';
    
    // Process paragraphs
    if (content.paragraphs) {
      content.paragraphs.forEach(para => {
        if (para.style?.startsWith('Heading')) {
          const level = para.style.replace('Heading', '');
          html += `<h${level}>${para.text}</h${level}>`;
        } else {
          let paraHtml = `<p>`;
          
          // Apply formatting if available
          if (para.formatting?.bold) paraHtml = `<strong>${paraHtml}`;
          if (para.formatting?.italic) paraHtml = `<em>${paraHtml}`;
          if (para.formatting?.underline) paraHtml = `<u>${paraHtml}`;
          
          paraHtml += para.text;
          
          // Close formatting tags
          if (para.formatting?.underline) paraHtml += `</u>`;
          if (para.formatting?.italic) paraHtml += `</em>`;
          if (para.formatting?.bold) paraHtml += `</strong>`;
          
          paraHtml += `</p>`;
          html += paraHtml;
        }
      });
    }
    
    // Process tables
    if (content.tables) {
      content.tables.forEach(table => {
        html += '<table class="fr-table">';
        
        table.rows.forEach(row => {
          html += '<tr>';
          row.cells.forEach(cell => {
            const tag = cell.isHeader ? 'th' : 'td';
            html += `<${tag}>${cell.text}</${tag}>`;
          });
          html += '</tr>';
        });
        
        html += '</table>';
      });
    }
    
    html += '</div>';
    return html;
  }
  
  // Convert Froala HTML back to structured content
  private convertHtmlToStructured(html: string): any {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    
    const structuredContent = {
      paragraphs: [],
      tables: []
    };
    
    // Process paragraphs and headings
    doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach(element => {
      const isHeading = element.tagName.toLowerCase().startsWith('h');
      const headingLevel = isHeading ? element.tagName.substring(1) : null;
      
      structuredContent.paragraphs.push({
        text: element.textContent,
        style: isHeading ? `Heading${headingLevel}` : 'Normal',
        formatting: {
          bold: element.querySelector('strong, b') !== null,
          italic: element.querySelector('em, i') !== null,
          underline: element.querySelector('u') !== null
        }
      });
    });
    
    // Process tables
    doc.querySelectorAll('table').forEach(tableElement => {
      const table = { rows: [] };
      
      tableElement.querySelectorAll('tr').forEach(row => {
        const cells = [];
        row.querySelectorAll('td, th').forEach(cell => {
          cells.push({
            text: cell.textContent,
            isHeader: cell.tagName.toLowerCase() === 'th'
          });
        });
        
        table.rows.push({ cells });
      });
      
      structuredContent.tables.push(table);
    });
    
    return structuredContent;
  }
}
```

### 3. DocxJS (for viewing DOCX in browser)
- **Link**: [DocxJS](https://github.com/VolodymyrBaydalka/docxjs)
- **Description**: JavaScript library for rendering DOCX directly in browser
- **Features**:
  - Direct DOCX rendering without conversion to HTML
  - High-fidelity document viewing
- **Use case**: When fidelity is more important than editing
- **Sample integration**:
```typescript
import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';
import { renderAsync } from 'docx-preview';

@Component({
  selector: 'app-docx-viewer',
  template: `<div #container></div>`,
  styles: []
})
export class DocxViewerComponent implements OnInit {
  @ViewChild('container', { static: true }) container: ElementRef;
  
  constructor() { }
  
  ngOnInit(): void { }
  
  async loadDocx(file: File) {
    const arrayBuffer = await file.arrayBuffer();
    await renderAsync(arrayBuffer, this.container.nativeElement);
  }
}
```

## Integrated Solution Architecture

For a complete .NET + Angular + MongoDB + Froala solution with DOCX support, I recommend this architecture:

### Backend (.NET API)
1. **DocumentController**: API endpoints for document operations
   ```csharp
   [ApiController]
   [Route("api/[controller]")]
   public class DocumentsController : ControllerBase
   {
       private readonly DocxMongoService _docxService;
       
       public DocumentsController(DocxMongoService docxService)
       {
           _docxService = docxService;
       }
       
       [HttpPost("import")]
       public async Task<IActionResult> ImportDocx([FromForm] IFormFile file, [FromForm] string title)
       {
           if (file == null || file.Length == 0)
               return BadRequest("No file uploaded");
               
           // Save uploaded file to temp location
           var tempFilePath = Path.GetTempFileName();
           using (var stream = new FileStream(tempFilePath, FileMode.Create))
           {
               await file.CopyToAsync(stream);
           }
           
           try
           {
               // Process and store in MongoDB
               var documentId = await _docxService.StoreDocxAsStructuredData(tempFilePath, title);
               return Ok(new { documentId });
           }
           finally
           {
               // Clean up temp file
               if (System.IO.File.Exists(tempFilePath))
                   System.IO.File.Delete(tempFilePath);
           }
       }
       
       [HttpGet("{id}")]
       public async Task<IActionResult> GetDocument(string id)
       {
           var document = await _docxService.GetDocumentById(id);
           if (document == null)
               return NotFound();
               
           return Ok(document);
       }
       
       [HttpPut("{id}")]
       public async Task<IActionResult> UpdateDocument(string id, DocumentUpdateModel model)
       {
           await _docxService.UpdateDocument(id, model.StructuredContent, model.HtmlCache);
           return Ok();
       }
       
       [HttpGet("{id}/download")]
       public async Task<IActionResult> DownloadAsDocx(string id)
       {
           var docxBytes = await _docxService.GenerateDocxFromStructured(id);
           return File(docxBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
       }
   }
   ```

2. **DocxMongoService**: Core service handling conversions and MongoDB integration
   - Implementation shown earlier in this document

### Frontend (Angular + Froala)
1. **Document Editor Component**:
   ```typescript
   import { Component, OnInit } from '@angular/core';
   import { ActivatedRoute } from '@angular/router';
   import { FroalaDocumentService } from './froala-document.service';
   
   @Component({
     selector: 'app-document-editor',
     template: `
       <div class="document-editor">
         <h1>{{ documentTitle }}</h1>
         
         <div class="editor-container">
           <div [froalaEditor]="froalaOptions" [(froalaModel)]="editorContent"></div>
         </div>
         
         <div class="actions">
           <button (click)="saveDocument()">Save Document</button>
           <button (click)="downloadAsDocx()">Download as DOCX</button>
         </div>
       </div>
     `,
     styles: []
   })
   export class DocumentEditorComponent implements OnInit {
     documentId: string;
     documentTitle: string = '';
     editorContent: string = '';
     
     froalaOptions: any = {
       // Froala options here
     };
     
     constructor(
       private route: ActivatedRoute,
       private documentService: FroalaDocumentService
     ) { }
     
     ngOnInit(): void {
       this.documentId = this.route.snapshot.paramMap.get('id');
       this.loadDocument();
     }
     
     loadDocument(): void {
       this.documentService.getDocument(this.documentId).subscribe(
         document => {
           this.documentTitle = document.title;
           this.editorContent = document.htmlForFroala;
         },
         error => console.error('Error loading document', error)
       );
     }
     
     saveDocument(): void {
       this.documentService.saveDocument(this.documentId, this.editorContent).subscribe(
         () => alert('Document saved successfully'),
         error => console.error('Error saving document', error)
       );
     }
     
     downloadAsDocx(): void {
       window.location.href = `https://your-api-domain.com/api/documents/${this.documentId}/download`;
     }
   }
   ```

2. **Document Import Component**:
   ```typescript
   import { Component } from '@angular/core';
   import { Router } from '@angular/router';
   import { FroalaDocumentService } from './froala-document.service';
   
   @Component({
     selector: 'app-document-import',
     template: `
       <div class="document-import">
         <h2>Import DOCX Document</h2>
         
         <div class="form-group">
           <label for="title">Document Title</label>
           <input type="text" id="title" [(ngModel)]="documentTitle">
         </div>
         
         <div class="form-group">
           <label for="file">Select DOCX File</label>
           <input type="file" id="file" (change)="onFileSelected($event)" accept=".docx">
         </div>
         
         <button [disabled]="!selectedFile || !documentTitle" (click)="uploadDocument()">
           Import Document
         </button>
       </div>
     `,
     styles: []
   })
   export class DocumentImportComponent {
     documentTitle: string = '';
     selectedFile: File = null;
     
     constructor(
       private documentService: FroalaDocumentService,
       private router: Router
     ) { }
     
     onFileSelected(event: any): void {
       this.selectedFile = event.target.files[0];
     }
     
     uploadDocument(): void {
       if (!this.selectedFile || !this.documentTitle) return;
       
       this.documentService.importDocx(this.selectedFile, this.documentTitle).subscribe(
         response => {
           // Navigate to editor with new document ID
           this.router.navigate(['/documents', response.documentId]);
         },
         error => console.error('Error importing document', error)
       );
     }
   }
   ```

## Recommendations

For your .NET + Angular + Froala + MongoDB stack with Option 3 approach, I recommend:

1. **Backend**:
   - Use Open XML SDK for .NET as the core library
   - Implement a DocxMongoService for handling conversions and DB operations
   - If budget allows, consider Aspose.Words for more robust handling

2. **Frontend**:
   - Use Mammoth.js for client-side conversion when needed
   - Implement a FroalaDocumentService to handle integration
   - Store HTML cache for direct loading in Froala

3. **Data Flow**:
   - Import: DOCX → Structured JSON + HTML Cache → MongoDB
   - Edit: MongoDB → HTML Cache → Froala
   - Save: Froala → Structured JSON + HTML Cache → MongoDB
   - Export: Structured JSON → DOCX

This approach gives you the best of both worlds - structured data for querying and manipulation, with HTML caching for optimal Froala performance.
