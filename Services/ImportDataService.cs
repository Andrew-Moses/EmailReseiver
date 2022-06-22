using System;
using System.Threading.Tasks;
using EmailReseiver.Contexts;
using EmailReseiver.Models;
using Microsoft.EntityFrameworkCore;

namespace EmailReseiver.MailServices
{
    public class ImportDataService
    {
        public ImportDataService(Context context)
        {
            _context = context;
        }
        public async Task<ImportData?> AddEntry(ImportData entry)
        {
            entry.InsertDate = DateTime.Now;
            await _context.AddAsync(entry);
            await _context.SaveChangesAsync();
            return await FindItem(entry.Id);
        }
 
        public Task<ImportData?> FindItem(int id) => 
            _context.ImportData.AsNoTracking()
                .FirstOrDefaultAsync(i => i.Id == id);
        private readonly Context _context;
    }
}