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
using Microsoft.AspNetCore.Http;

namespace ProjectDipMVC.Controllers
{
    [Authorize(Roles = "Администратор, Редактор, Главный редактор")]
    public class SectionsProjectsController : Controller
    {
        private readonly ProjectDipContext _context;

        public SectionsProjectsController(ProjectDipContext context)
        {
            _context = context;
        }

        // GET: SectionsProjects
        public async Task<IActionResult> Index()
        {
            var UserId = 1;
            var projectDipContext =
            from pd in _context.ProjectDescripts.Include(s => s.Project)
            join p in _context.SectionsProjects on 
                pd.ProjDscrptId equals p.ProjDscrptId into ps
            from p in ps.DefaultIfEmpty()
            where pd.UserId == UserId
            select new SectionsProjectIndex { 
                ProjDscrptId = pd.ProjDscrptId,
                Name = pd.Project.Name,
                Section_Name = pd.SectionName,
                Section_Number = pd.SectionNumber,
                SectionsId = null != p ? p.SectionsId: null,
                NameSections = null != p ? p.NameSections : null,
                NumberSections = null != p ? p.NumberSections : null,
                NameFileSections = null != p ? p.NameFileSections : null
            };

            //var projectDipContext = _context.SectionsProjects.Include(s => s.ProjDscrpt);//.Where(p => p.);
            return View(await projectDipContext.ToListAsync());
        }

        // GET: SectionsProjects/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null || _context.SectionsProjects == null)
            {
                return NotFound();
            }

            var sectionsProject = await _context.SectionsProjects
                .Include(s => s.ProjDscrpt)
                .FirstOrDefaultAsync(m => m.SectionsId == id);
            if (sectionsProject == null)
            {
                return NotFound();
            }

            return View(sectionsProject);
        }

        // GET: SectionsProjects/Create
        public IActionResult Create(int? ProjDscrptId)
        {
            if (null == ProjDscrptId)
            {
                return NotFound();
            }

            if (null == _context.ProjectDescripts.Find(ProjDscrptId))
            {
                return NotFound();
            }

            ViewBag.ProjDscrptId = ProjDscrptId;
            ViewData["ProjDscrptId"] = new SelectList(_context.ProjectDescripts, "ProjDscrptId", "ProjDscrptId");

            ViewBag.ProjDscrpt = _context.ProjectDescripts;
            return View();
        }

        // POST: SectionsProjects/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("SectionsId,NameSections,NumberSections,ProjDscrptId,NameFileSections,FileSections")] SectionsProject sectionsProject, IFormFile FileSections)
        {

            if (null == _context.ProjectDescripts.Find(sectionsProject.ProjDscrptId))
            {
                return NotFound();
            }

            ViewData["ProjDscrptId"] = new SelectList(_context.ProjectDescripts, "ProjDscrptId", "ProjDscrptId", sectionsProject.ProjDscrptId);
            ViewBag.ProjDscrpt = _context.ProjectDescripts;

            bool isWrong = false;

            if (String.IsNullOrEmpty(sectionsProject.NameSections))
            {
                ModelState.AddModelError("NameSections", "Введите название секции");
                isWrong = true;
            }

            if ((null == sectionsProject.NumberSections) || (sectionsProject.NumberSections <= 0))
            {
                ModelState.AddModelError("NumberSections", "Введите корректный номер секции");
                isWrong = true;
            }

            if (null == (FileSections))
            {
                ModelState.AddModelError("FileSections", "Выберите файл");
                isWrong = true;
            }

            if (isWrong)
            {
                return View(sectionsProject);
            }

            sectionsProject.NameFileSections = FileSections.FileName;
            using (var ms = new MemoryStream())
            {
                FileSections.CopyTo(ms);
                sectionsProject.FileSections = ms.ToArray();
            }

            _context.Add(sectionsProject);
            try
            {
                await _context.SaveChangesAsync();
            }
            catch (DbUpdateException)
            {
                // ModelState.AddModelError("ProjDscrptId", "Секция с таким заданием уже занята");
                TempData["ErrMessage"] = "Секция с таким заданием уже занята";
                return View(sectionsProject);
            }
            return RedirectToAction(nameof(Index));
        }

        // GET: SectionsProjects/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null || _context.SectionsProjects == null)
            {
                return NotFound();
            }

            if (null == _context.SectionsProjects.Find(id))
            {
                return NotFound();
            }

            var sectionsProject = await _context.SectionsProjects.FindAsync(id);
            if (sectionsProject == null)
            {
                return NotFound();
            }
            ViewData["ProjDscrptId"] = new SelectList(_context.ProjectDescripts, "ProjDscrptId", "ProjDscrptId", sectionsProject.ProjDscrptId);
            ViewBag.ProjDscrpt = _context.ProjectDescripts;
            return View(sectionsProject);
        }

        // POST: SectionsProjects/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("SectionsId,NameSections,NumberSections,ProjDscrptId,NameFileSections,FileSections")] SectionsProject sectionsProject, IFormFile FileSections)
        {
            if (id != sectionsProject.SectionsId)
            {
                return NotFound();
            }

            if (null == _context.ProjectDescripts.Find(sectionsProject.ProjDscrptId))
            {
                return NotFound();
            }

            ViewData["ProjDscrptId"] = new SelectList(_context.ProjectDescripts, "ProjDscrptId", "ProjDscrptId", sectionsProject.ProjDscrptId);
            ViewBag.ProjDscrpt = _context.ProjectDescripts;

            bool isWrong = false;

            if (String.IsNullOrEmpty(sectionsProject.NameSections))
            {
                ModelState.AddModelError("NameSections", "Введите название секции");
                isWrong = true;
            }

            if ((null == sectionsProject.NumberSections) || (sectionsProject.NumberSections <= 0))
            {
                ModelState.AddModelError("NumberSections", "Введите корректный номер секции");
                isWrong = true;
            }

            if (null == (FileSections))
            {
                ModelState.AddModelError("FileSections", "Выберите файл");
                isWrong = true;
            }

            if (isWrong)
            {
                return View(sectionsProject);
            }

            if (ModelState.IsValid)
            {
                try
                {

                    sectionsProject.NameFileSections = FileSections.FileName;
                    using (var ms = new MemoryStream())
                    {
                        FileSections.CopyTo(ms);
                        sectionsProject.FileSections = ms.ToArray();
                    }

                    _context.Update(sectionsProject);
                    try
                    {
                        await _context.SaveChangesAsync();
                    }
                    catch (DbUpdateException)
                    {
                        TempData["ErrMessage"] = "Секция с таким заданием уже занята";
                        return View(sectionsProject);
                    }
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!SectionsProjectExists(sectionsProject.SectionsId))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            ViewData["ProjDscrptId"] = new SelectList(_context.ProjectDescripts, "ProjDscrptId", "ProjDscrptId", sectionsProject.ProjDscrptId);
            return View(sectionsProject);
        }

        // GET: SectionsProjects/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null || _context.SectionsProjects == null)
            {
                return NotFound();
            }

            var sectionsProject = await _context.SectionsProjects
                .Include(s => s.ProjDscrpt)
                .FirstOrDefaultAsync(m => m.SectionsId == id);
            if (sectionsProject == null)
            {
                return NotFound();
            }

            return View(sectionsProject);
        }

        // POST: SectionsProjects/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            if (_context.SectionsProjects == null)
            {
                return Problem("Entity set 'ProjectDipContext.SectionsProjects'  is null.");
            }
            var sectionsProject = await _context.SectionsProjects.FindAsync(id);
            if (sectionsProject != null)
            {
                _context.SectionsProjects.Remove(sectionsProject);
            }
            
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool SectionsProjectExists(int id)
        {
          return (_context.SectionsProjects?.Any(e => e.SectionsId == id)).GetValueOrDefault();
        }

    }
}
