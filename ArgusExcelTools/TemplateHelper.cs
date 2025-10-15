using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ArgusExcelTools
{
    public static class TemplateHelper
    {
        public static string ExtractTemplate(string resourcePartialName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();

            // Find the embedded resource that ends with the given name
            string fullName = asm.GetManifestResourceNames()
                                 .FirstOrDefault(n => n.EndsWith(resourcePartialName, StringComparison.OrdinalIgnoreCase));
            if (fullName == null)
                throw new FileNotFoundException($"Embedded resource not found: {resourcePartialName}");

            // Create a temp file with the same extension
            string tempFile = Path.Combine(Path.GetTempPath(), resourcePartialName);

            using (Stream rs = asm.GetManifestResourceStream(fullName))
            using (FileStream fs = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
            {
                rs.CopyTo(fs);
            }

            return tempFile;
        }
    }
}
