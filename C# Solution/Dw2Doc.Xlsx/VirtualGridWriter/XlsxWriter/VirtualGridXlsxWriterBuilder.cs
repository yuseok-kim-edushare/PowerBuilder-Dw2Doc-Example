﻿using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGrid;
using Appeon.DotnetDemo.Dw2Doc.Common.VirtualGridWriter.Abstractions;

namespace Appeon.DotnetDemo.Dw2Doc.Xlsx.VirtualGridWriter.XlsxWriter
{
    public class VirtualGridXlsxWriterBuilder : IVirtualGridWriterBuilder
    {
        public bool Append { get; set; }
        public string? WriteToPath { get; set; }
        public string? AppendToSheetName { get; set; }

        public AbstractVirtualGridWriter? Build(VirtualGrid grid,
            out string? error)
        {
            error = null;
            if (WriteToPath is not null)
            {
                try
                {
                    if (Append)
                    {
                        if (!File.Exists(WriteToPath))
                        {
                            error = $"Specified path [{WriteToPath}] does not exist";
                            return null;
                        }

                        if (AppendToSheetName is null)
                        {
                            error = "Did not specify a sheet name";
                            return null;
                        }
                        return new VirtualGridXlsxWriter(grid, WriteToPath, AppendToSheetName);
                    }
                    return new VirtualGridXlsxWriter(grid, WriteToPath);
                }
                catch (Exception e)
                {
                    error = e.Message;
                    return null;
                }
            }

            error = "Must specify a path";

            return null;
        }

        public AbstractVirtualGridWriter? BuildFromTemplate(VirtualGrid grid,
            string filePath,
            out string? error)
        {
            error = null;
            throw new NotImplementedException();
            // Future implementation will go here
        }
    }
}
