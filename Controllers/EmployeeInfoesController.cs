using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using CRUD.MVC.WebApp.Models;

namespace CRUD.MVC.WebApp.Controllers
{
    public class EmployeeInfoesController : Controller
    {
        private CRUDEntities db = new CRUDEntities();

        // GET: EmployeeInfoes
        public ActionResult Index()
        {
            return View(db.EmployeeInfoes.ToList());
        }

        // GET: EmployeeInfoes/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EmployeeInfo employeeInfo = db.EmployeeInfoes.Find(id);
            if (employeeInfo == null)
            {
                return HttpNotFound();
            }
            return View(employeeInfo);
        }

        // GET: EmployeeInfoes/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: EmployeeInfoes/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EmployeeId,EmployeeName,Address,IsActive")] EmployeeInfo employeeInfo)
        {
            employeeInfo.CreationDate = DateTime.Now;
            if (ModelState.IsValid)
            {
                db.EmployeeInfoes.Add(employeeInfo);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(employeeInfo);
        }

        // GET: EmployeeInfoes/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EmployeeInfo employeeInfo = db.EmployeeInfoes.Find(id);
            if (employeeInfo == null)
            {
                return HttpNotFound();
            }
            return View(employeeInfo);
        }

        // POST: EmployeeInfoes/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EmployeeId,EmployeeName,Address,IsActive,CreationDate")] EmployeeInfo employeeInfo)
        {
            if (ModelState.IsValid)
            {
                db.Entry(employeeInfo).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(employeeInfo);
        }

        // GET: EmployeeInfoes/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            EmployeeInfo employeeInfo = db.EmployeeInfoes.Find(id);
            if (employeeInfo == null)
            {
                return HttpNotFound();
            }
            return View(employeeInfo);
        }

        // POST: EmployeeInfoes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            EmployeeInfo employeeInfo = db.EmployeeInfoes.Find(id);
            db.EmployeeInfoes.Remove(employeeInfo);
            db.SaveChanges();
            return RedirectToAction("Index");
        }


        public ActionResult Export()
        {
            var employees = db.EmployeeInfoes.ToList();
            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            string fileName = "Employees.xlsx";
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    IXLWorksheet worksheet =
                    workbook.Worksheets.Add("Employees");
                    worksheet.Cell(1, 1).Value = "EmployeeId";
                    worksheet.Cell(1, 2).Value = "EmployeeName";
                    worksheet.Cell(1, 3).Value = "Address";
                    worksheet.Cell(1, 4).Value = "IsActive";
                    worksheet.Cell(1, 5).Value = "CreationDate";
                    for (int index = 1; index <= employees.Count; index++)
                    {
                        worksheet.Cell(index + 1, 1).Value = employees[index - 1].EmployeeId;
                        worksheet.Cell(index + 1, 2).Value = employees[index - 1].EmployeeName;
                        worksheet.Cell(index + 1, 3).Value = employees[index - 1].Address;
                        worksheet.Cell(index + 1, 4).Value = employees[index - 1].IsActive;
                        worksheet.Cell(index + 1, 5).Value = employees[index - 1].CreationDate;
                    }
                    //required using System.IO;
                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, contentType, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = ex.Message;
                return View("Index");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
