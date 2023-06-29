﻿using ClosedXML.Excel;
using ClosedXML.Extensions;
using Aspose.Cells;
using DotNetApi.Infrastructure;
using DotNetApi.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Data;
using DotNetApi.Controllers;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace TodoApi.Controllers;


[Route("api/[controller]")]
[ApiController]
public class TodoItemsController : ControllerBase
{
    private readonly DataDbContext _context;
    //private readonly IExportExcelService _exportExcelService;
    //private List<object> todoItems;

    //public TodoItemsController(DataDbContext context, IExportExcelService exportExcelService)
    public TodoItemsController(DataDbContext context)
    {
        _context = context;
        //_exportExcelService = exportExcelService;
    }

    // GET: api/TodoItems
    [HttpGet]
    public async Task<ActionResult<IEnumerable<TodoItemDTO>>> GetTodoItems()
    {
        return await _context.TodoItems
            .Select(x => ItemToDTO(x))
            .ToListAsync();
    }

    // GET: api/TodoItems/5
    // <snippet_GetByID>
    [HttpGet("{id}")]
    public async Task<ActionResult<TodoItemDTO>> GetTodoItem(int id)
    {
        var todoItem = await _context.TodoItems.FindAsync(id);

        if (todoItem == null)
        {
            return NotFound();
        }

        return ItemToDTO(todoItem);
    }
    // </snippet_GetByID>

    // PUT: api/TodoItems/5
    // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
    // <snippet_Update>
    [HttpPut("{id}")]
    public async Task<IActionResult> PutTodoItem(int id, TodoItemDTO todoDTO)
    {
        if (id != todoDTO.Id)
        {
            return BadRequest();
        }

        var todoItem = await _context.TodoItems.FindAsync(id);
        if (todoItem == null)
        {
            return NotFound();
        }

        todoItem.Name = todoDTO.Name;
        todoItem.IsComplete = todoDTO.IsComplete;

        try
        {
            await _context.SaveChangesAsync();
        }
        catch (DbUpdateConcurrencyException) when (!TodoItemExists(id))
        {
            return NotFound();
        }

        return NoContent();
    }
    // </snippet_Update>

    // POST: api/TodoItems
    // To protect from overposting attacks, see https://go.microsoft.com/fwlink/?linkid=2123754
    // <snippet_Create>
    [HttpPost]
    public async Task<ActionResult<TodoItemDTO>> PostTodoItem(TodoItemDTO todoDTO)
    {
        var todoItem = new TodoItem
        {
            IsComplete = todoDTO.IsComplete,
            Name = todoDTO.Name
        };

        _context.TodoItems.Add(todoItem);
        await _context.SaveChangesAsync();

        return CreatedAtAction(
            nameof(GetTodoItem),
            new { id = todoItem.Id },
            ItemToDTO(todoItem));
    }
    // </snippet_Create>

    // DELETE: api/TodoItems/5
    [HttpDelete("{id}")]
    public async Task<IActionResult> DeleteTodoItem(int id)
    {
        var todoItem = await _context.TodoItems.FindAsync(id);
        if (todoItem == null)
        {
            return NotFound();
        }

        _context.TodoItems.Remove(todoItem);
        await _context.SaveChangesAsync();

        return NoContent();
    }

    private bool TodoItemExists(int id)
    {
        return _context.TodoItems.Any(e => e.Id == id);
    }

    [HttpGet(template:"ExportExcel")]
    public async Task<ActionResult> GetExcel()
    {
        ExportExcelService exportExcelService = new ExportExcelService();
        List<TodoItem> todoItems = new List<TodoItem>();
        DataTable dt = exportExcelService.TodoItemTable(todoItems);
        XLWorkbook wb = exportExcelService.ExportExcel(dt, title: "Лабораторная № 1", fileName: "");
        using (MemoryStream ms = new())
        {
            wb.SaveAs(ms);
            return File(fileContents: ms.ToArray(), contentType: "application/vnd.openxmlformats-officedocument.spreadsheetsml.sheet", fileDownloadName: "Labs_TodoItem"); 
        }

    }

    private static TodoItemDTO ItemToDTO(TodoItem todoItem) =>
       new TodoItemDTO
       {
           Id = todoItem.Id,
           Name = todoItem.Name,
           IsComplete = todoItem.IsComplete
       };
    
}