using System;
using System.Threading.Tasks;
using EmailReseiver.Contexts;
using EmailReseiver.Models;
using Microsoft.EntityFrameworkCore;

namespace EmailReseiver.Services
{
    public class LogService
    {
        public LogService(Context context)
        {
            _context = context;
        }

        public async Task<ExceptionLogImportData?> AddEntry(ExceptionLogImportData entry)
        {
            entry.LoggedDate = DateTime.Now;
            await _context.AddAsync(entry);
            await _context.SaveChangesAsync();
            return await FindItem(entry.Id);
        }

        public Task<ExceptionLogImportData?> FindItem(Int64 id) =>
            _context.ExceptionLogImportData.AsNoTracking()
                .FirstOrDefaultAsync(i => i.Id == id);
        private readonly Context _context;
    }
}
