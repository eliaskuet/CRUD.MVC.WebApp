using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
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
