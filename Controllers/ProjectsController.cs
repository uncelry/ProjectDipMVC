using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using ProjectDipMVC.Models;
using Microsoft.AspNetCore.Authorization;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using Syncfusion.OfficeChart;

namespace ProjectDipMVC.Controllers
{

    [Authorize(Roles = "Администратор, Гланый редактор")]
    public class ProjectsController : Controller
    {
        private readonly ProjectDipContext _context;
        IWebHostEnvironment _appEnvironment;

        public ProjectsController(ProjectDipContext context, IWebHostEnvironment appEnvironment)
        {
            _context = context;
            _appEnvironment = appEnvironment;
        }

        // GET: Projects
        public async Task<IActionResult> Index()
        {
            var projectDipContext = _context.Projects.Include(p => p.User);
            try
            {
                var t = projectDipContext.ToList();
            }catch(Exception ex)
            {
                var t = ex;
            }
                return View(await projectDipContext.ToListAsync());
        }

        // GET: Projects/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null || _context.Projects == null)
            {
                return NotFound();
            }

            var project = await _context.Projects
                .Include(p => p.User)
                .FirstOrDefaultAsync(m => m.ProjectId == id);
            if (project == null)
            {
                return NotFound();
            }

            return View(project);
        }

        // GET: Projects/Create
        public IActionResult Create()
        {
            ViewData["UserId"] = new SelectList(_context.Users, "UserId", "UserName");
            return View();
        }

        // POST: Projects/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ProjectUpload projectUpload)
        {

            bool isWrong = false;

            ViewData["UserId"] = new SelectList(_context.Users, "UserId", "UserName");
            if (String.IsNullOrEmpty(projectUpload.Name))
            {
                ModelState.AddModelError("Name", "Введите название проекта");
                isWrong = true;
            }

            if (null == (_context.Users.Find(projectUpload.UserId)))
            {
                ModelState.AddModelError("UserId", "Выберите редактора");
                isWrong = true;
            }

            if (null == (projectUpload.TitulFile))
            {
                ModelState.AddModelError("TitulFile", "Выберите файл");
                isWrong = true;
            }

            if (isWrong)
            {
                return View(projectUpload);
            }

            var project = createProject(projectUpload);
            _context.Add(project);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        // GET: Projects/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null || _context.Projects == null)
            {
                return NotFound();
            }

            var project = await _context.Projects.FindAsync(id);
            if (project == null)
            {
                return NotFound();
            }
            var projectUpload = new ProjectUpload
            {
                ProjectId = project.ProjectId,
                Name = project.Name,
                UserId = project.UserId
            };
            ViewData["UserId"] = new SelectList(_context.Users, "UserId", "UserName", projectUpload.UserId);
            
            return View(projectUpload);
        }

        // POST: Projects/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ProjectUpload projectUpload)
        {

            if (id != projectUpload.ProjectId)
            {
                return NotFound();
            }

            bool isWrong = false;

            ViewData["UserId"] = new SelectList(_context.Users, "UserId", "UserName");
            if (String.IsNullOrEmpty(projectUpload.Name))
            {
                ModelState.AddModelError("Name", "Введите название проекта");
                isWrong = true;
            }

            if (null == (_context.Users.Find(projectUpload.UserId)))
            {
                ModelState.AddModelError("UserId", "Выберите редактора");
                isWrong = true;
            }

            if (null == (projectUpload.TitulFile))
            {
                ModelState.AddModelError("TitulFile", "Выберите файл");
                isWrong = true;
            }

            if (isWrong)
            {
                return View(projectUpload);
            }

            try
            {
                var project = createProject(projectUpload, id);
                _context.Update(project);
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!ProjectExists(projectUpload.ProjectId))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }
            return RedirectToAction(nameof(Index));

            ViewData["UserId"] = new SelectList(_context.Users, "UserId", "UserName", projectUpload.UserId);
            return View(projectUpload);
        }

        // GET: Projects/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null || _context.Projects == null)
            {
                return NotFound();
            }

            var project = await _context.Projects
                .Include(p => p.User)
                .FirstOrDefaultAsync(m => m.ProjectId == id);
            if (project == null)
            {
                return NotFound();
            }

            return View(project);
        }

        // POST: Projects/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            if (_context.Projects == null)
            {
                return Problem("Ошибка");
            }

            

            var project = await _context.Projects.FindAsync(id);
            if (project != null)
            {
                _context.Projects.Remove(project);
            }

            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateException)
            {
                project = await _context.Projects
                .Include(p => p.User)
                .FirstOrDefaultAsync(m => m.ProjectId == id);
                TempData["ErrMessage"] = "Удаление невозможно - убедитесь, что отсутствуют задачи, связанные с данным проектом";
                return View(project);
            }
            return RedirectToAction(nameof(Index));
        }

        private bool ProjectExists(int id)
        {
          return (_context.Projects?.Any(e => e.ProjectId == id)).GetValueOrDefault();
        }

        public async Task<IActionResult> Build(int? id)
        {
            if (id == null || _context.Projects == null)
            {
                return NotFound();
            }
            // Получаем проект
            var project = await _context.Projects.FindAsync(id);

            // Получаем задания по проекту
            var descripts = _context.ProjectDescripts.Where(p => p.ProjectId == project.ProjectId);

            // Список под секции
            var sectionList = new List<SectionsProject>();

            // Получаем все секции, связанные с проектом
            foreach(var descript in descripts)
            {
                
                sectionList.AddRange(_context.SectionsProjects.Where(p => p.ProjDscrptId == descript.ProjDscrptId));
            }

            // Сортируем
            sectionList.Sort((x, y) => ((int)x.ProjDscrptId - (int)y.ProjDscrptId));

            if (project == null)
            {
                return NotFound();
            }
            var titulFile = project.TitulFile;

            MemoryStream memStream = new MemoryStream(titulFile);
            IFormFile uploadedFile = new FormFile(memStream, 0, memStream.Length, project.TitulName, project.TitulName);

            // Создаем локальные файлы для каждой секции
            foreach(var section in sectionList)
            {
                MemoryStream tmpStream = new MemoryStream(section.FileSections);
                IFormFile tmpFile = new FormFile(tmpStream, 0, tmpStream.Length, section.NameFileSections, section.NameFileSections);
                string tmpPath = "/Files/" + tmpFile.FileName;
                using (var fileStream = new FileStream(_appEnvironment.WebRootPath + tmpPath, FileMode.Create))
                {
                    await tmpFile.CopyToAsync(fileStream);
                }
            }

            // путь к папке Files
            string path = "/Files/" + uploadedFile.FileName;
            // сохраняем файл в папку Files в каталоге wwwroot
            using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
            {
                await uploadedFile.CopyToAsync(fileStream);
            }

            // Список под файлы секций
            var sourcesList = new List<FileStream>();
            foreach(var section in sectionList)
            {
                string tmpPath = "/Files/" + section.NameFileSections;
                sourcesList.Add(new FileStream(_appEnvironment.WebRootPath + tmpPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
            }

            FileStream destinationStreamPath = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);


            // Добавляем к титульнику раздел содержания
            // Открываем документ назначения
            WordDocument destinationDocument = new WordDocument(destinationStreamPath, FormatType.Docx);
            MemoryStream stream = new MemoryStream();

            // Добавляем секцию в документ
            IWSection sectionTOC = destinationDocument.AddSection();

            // string paraText = "Оглавление";

            // Добавляем абзац в секцию (для заголовка)
            IWParagraph paragraphName = sectionTOC.AddParagraph();
            paragraphName.AppendText("Оглавление");

            IWParagraphStyle style = destinationDocument.AddParagraphStyle("TOC_Name");
            style.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            style.CharacterFormat.FontName = "Times New Roman";
            style.CharacterFormat.Bold = true;
            style.CharacterFormat.FontSize = 16f;
            paragraphName.ApplyStyle("TOC_Name");


            // Добавляем абзац в секцию
            IWParagraph paragraph = sectionTOC.AddParagraph();

            WParagraphStyle styleTOC = (WParagraphStyle)destinationDocument.AddParagraphStyle("Mystyle");
            styleTOC.CharacterFormat.FontName = "Times New Roman";
            styleTOC.CharacterFormat.Bold = false;
            styleTOC.CharacterFormat.FontSize = 14f;

            IWCharacterStyle cStyleToc = destinationDocument.AddCharacterStyle("Mystyle");
            styleTOC.CharacterFormat.FontName = "Times New Roman";
            styleTOC.CharacterFormat.Bold = false;
            styleTOC.CharacterFormat.FontSize = 14f;



            //TableOfContent tableOfContents = new TableOfContent(destinationDocument, "\\p");
            // Добавляем оглавление в абзац
            paragraph.ApplyStyle("Mystyle");

            TableOfContent tableOfContents = paragraph.AppendTOC(1, 3);
            tableOfContents.LowerHeadingLevel = 1;
            tableOfContents.UpperHeadingLevel = 3;

            paragraph.AppendBreak(BreakType.PageBreak);
            paragraph.ParagraphFormat.LeftIndent = 11;
            paragraph.ParagraphFormat.RightIndent = 11;
            tableOfContents.UseOutlineLevels = true;
            tableOfContents.SetTOCLevelStyle(1, "Mystyle");
            tableOfContents.UseTableEntryFields = true;
            tableOfContents.ApplyStyle("Mystyle");

            // Сохраняем файл
            destinationDocument.Save(stream, FormatType.Docx);

            // Объединяем секции и титульник
            // Проходим по каждой секции
            foreach (var src in sourcesList)
            {
                // Открываем секцию в виде WordDocument
                using (WordDocument document = new WordDocument(src, FormatType.Automatic))
                {
                    // Импортируем содержимое документа источника в конец документа назначения
                    destinationDocument.ImportContent(document, ImportOptions.UseDestinationStyles);

                    // Сохраняем и закрываем источник в MemoryStream
                    destinationDocument.Save(stream, FormatType.Docx);
                    document.Close();
                }
            }

            // Обновляем оглавление
            destinationDocument.UpdateTableOfContents();

            

            // Сохраняем файл
            destinationDocument.Save(stream, FormatType.Docx);

            DocIORenderer render = new DocIORenderer();
            render.Settings.ChartRenderingOptions.ImageFormat = ExportImageFormat.Jpeg;
            PdfDocument pdfDocument = render.ConvertToPDF(destinationDocument);
            render.Dispose();
            destinationDocument.Dispose();

            MemoryStream outputStream = new MemoryStream();
            pdfDocument.Save(outputStream);
            outputStream.Position = 0;
            pdfDocument.Close();

            // Закрываем файл
            destinationDocument.Close();
            stream.Position = 0;


            // Закрываем поток титульника и удаляем его
            destinationStreamPath.Close();
            System.IO.File.Delete(_appEnvironment.WebRootPath + path);

            // Закрываем потоки секций
            foreach (var secFile in sourcesList)
            {
                secFile.Close();
            }

            // Удаляем секции
            foreach (var section in sectionList)
            {
                string tmpPath = "/Files/" + section.NameFileSections;
                System.IO.File.Delete(_appEnvironment.WebRootPath + tmpPath);
            }


            //return File(stream, "application/msword", "Result.docx");
            FileStreamResult fileStreamResult = new FileStreamResult(outputStream, "application/pdf");
            fileStreamResult.FileDownloadName = "Sample.pdf";

            return fileStreamResult;
            //return NoContent();
        }

        private static Project createProject(ProjectUpload projectUpload, int? ProjectId = null)
        {
            var fileName = Path.GetFileName(projectUpload.TitulFile.FileName);
            var project = new Project
            {
                Name = projectUpload.Name,
                DateCreate = DateTime.Now,
                UserId = projectUpload.UserId,
                TitulName = fileName
            };
            project.ProjectId = ProjectId != null ? ProjectId.Value : project.ProjectId;
            using (var target = new MemoryStream())
            {
                projectUpload.TitulFile.CopyTo(target);
                project.TitulFile = target.ToArray();
            }

            return project;
        }
    }
}
