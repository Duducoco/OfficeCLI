// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.CommandLine;
using System.Text;
using OfficeCli.Core;

namespace OfficeCli;

static partial class CommandBuilder
{
    private static Command BuildImportCommand(Option<bool> jsonOption)
    {
        var importFileArg = new Argument<FileInfo>("file") { Description = "Target Office file (.xlsx or .docx)" };
        var importParentPathArg = new Argument<string>("parent-path") { Description = "Target path: Excel sheet (e.g. /Sheet1) or Word body (/body)" };
        var importSourceArg = new Argument<FileInfo?>("source-file") { Description = "Source data file to import (positional, alternative to --file)" };
        importSourceArg.DefaultValueFactory = _ => null!;
        var importSourceOpt = new Option<FileInfo?>("--file") { Description = "Source data file to import" };
        var importStdinOpt = new Option<bool>("--stdin") { Description = "Read data from stdin" };
        var importFormatOpt = new Option<string?>("--format") { Description = "Data format: csv/tsv (xlsx) or markdown/md (docx); default inferred from source extension" };
        var importHeaderOpt = new Option<bool>("--header") { Description = "First row is header: set AutoFilter and freeze pane" };
        var importStartCellOpt = new Option<string>("--start-cell") { Description = "Starting cell (default: A1)" };
        var importStyleSourceOpt = new Option<FileInfo?>("--style-source") { Description = "For docx markdown import: extract heading/body style mapping from this .docx" };
        importStartCellOpt.DefaultValueFactory = _ => "A1";

        var importCommand = new Command("import", "Import data into an Office file (xlsx: CSV/TSV, docx: Markdown)");
        importCommand.Add(importFileArg);
        importCommand.Add(importParentPathArg);
        importCommand.Add(importSourceArg);
        importCommand.Add(importSourceOpt);
        importCommand.Add(importStdinOpt);
        importCommand.Add(importFormatOpt);
        importCommand.Add(importHeaderOpt);
        importCommand.Add(importStartCellOpt);
        importCommand.Add(importStyleSourceOpt);
        importCommand.Add(jsonOption);

        importCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(importFileArg)!;
            var parentPath = result.GetValue(importParentPathArg)!;
            var source = result.GetValue(importSourceOpt) ?? result.GetValue(importSourceArg);
            var useStdin = result.GetValue(importStdinOpt);
            var format = result.GetValue(importFormatOpt);
            var header = result.GetValue(importHeaderOpt);
            var startCell = result.GetValue(importStartCellOpt)!;
            var styleSource = result.GetValue(importStyleSourceOpt);

            if (!file.Exists)
                throw new CliException($"File not found: {file.FullName}")
                {
                    Code = "file_not_found",
                    Suggestion = $"Create the file first: officecli create \"{file.FullName}\""
                };

            var ext = Path.GetExtension(file.FullName).ToLowerInvariant();
            if (ext is not ".xlsx" and not ".docx")
                throw new CliException("Import currently supports .xlsx (csv/tsv) and .docx (markdown)")
                {
                    Code = "unsupported_type",
                    Suggestion = "Use a .xlsx or .docx file"
                };

            // Read source content
            string sourceContent;
            if (useStdin)
            {
                sourceContent = Console.In.ReadToEnd();
            }
            else if (source != null)
            {
                if (!source.Exists)
                    throw new CliException($"Source file not found: {source.FullName}")
                    {
                        Code = "file_not_found"
                    };
                sourceContent = File.ReadAllText(source.FullName, Encoding.UTF8);
            }
            else
            {
                throw new CliException("Either --file or --stdin must be specified")
                {
                    Code = "missing_argument",
                    Suggestion = "Use --file <path> to specify a source file, or --stdin to read from standard input"
                };
            }

            if (ext == ".xlsx")
            {
                // Determine delimiter: --format flag > source file extension > default csv
                char delimiter = ',';
                if (!string.IsNullOrEmpty(format))
                {
                    delimiter = format.ToLowerInvariant() switch
                    {
                        "tsv" => '\t',
                        "csv" => ',',
                        _ => throw new CliException($"Unknown format: {format}. Use 'csv' or 'tsv' for .xlsx import")
                        {
                            Code = "invalid_value",
                            ValidValues = ["csv", "tsv"]
                        }
                    };
                }
                else if (source != null)
                {
                    var sourceExt = Path.GetExtension(source.FullName).ToLowerInvariant();
                    if (sourceExt == ".tsv" || sourceExt == ".tab")
                        delimiter = '\t';
                }

                using var handler = new OfficeCli.Handlers.ExcelHandler(file.FullName, editable: true);
                var msg = handler.Import(parentPath, sourceContent, delimiter, header, startCell);
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else
                    Console.WriteLine(msg);
            }
            else
            {
                var fmt = format?.ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(fmt) && source != null)
                {
                    var sourceExt = Path.GetExtension(source.FullName).ToLowerInvariant();
                    if (sourceExt is ".md" or ".markdown")
                        fmt = "markdown";
                }
                fmt ??= "markdown";

                if (fmt is not "markdown" and not "md")
                    throw new CliException($"Unknown format: {format}. Use 'markdown' or 'md' for .docx import")
                    {
                        Code = "invalid_value",
                        ValidValues = ["markdown", "md"]
                    };

                if (styleSource != null)
                {
                    if (!styleSource.Exists)
                        throw new CliException($"Style source file not found: {styleSource.FullName}")
                        {
                            Code = "file_not_found"
                        };
                    if (!string.Equals(Path.GetExtension(styleSource.FullName), ".docx", StringComparison.OrdinalIgnoreCase))
                        throw new CliException("--style-source must be a .docx file")
                        {
                            Code = "invalid_value"
                        };
                }

                using var handler = new OfficeCli.Handlers.WordHandler(file.FullName, editable: true);
                var msg = handler.ImportMarkdown(parentPath, sourceContent, styleSource?.FullName);
                if (json)
                    Console.WriteLine(OutputFormatter.WrapEnvelopeText(msg));
                else
                    Console.WriteLine(msg);
            }
            return 0;
        }, json); });

        return importCommand;
    }

    private static Command BuildCreateCommand(Option<bool> jsonOption)
    {
        var createFileArg = new Argument<string>("file") { Description = "Output file path (.docx, .xlsx, .pptx)" };
        var createTypeOpt = new Option<string>("--type") { Description = "Document type (docx, xlsx, pptx) — optional, inferred from file extension" };
        var createCommand = new Command("create", "Create a blank Office document");
        createCommand.Aliases.Add("new");
        createCommand.Add(createFileArg);
        createCommand.Add(createTypeOpt);
        createCommand.Add(jsonOption);

        createCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var file = result.GetValue(createFileArg)!;
            var type = result.GetValue(createTypeOpt);

            // If file has no extension but --type is provided, append it
            if (!string.IsNullOrEmpty(type) && string.IsNullOrEmpty(Path.GetExtension(file)))
            {
                var ext = type.StartsWith('.') ? type : "." + type;
                file += ext;
            }

            // Check if the file is held by a resident process
            var fullPath = Path.GetFullPath(file);
            if (ResidentClient.TryConnect(fullPath, out _))
            {
                throw new CliException($"{Path.GetFileName(file)} is currently opened by a resident process. Please run 'officecli close \"{file}\"' first.")
                {
                    Code = "file_locked",
                    Suggestion = $"Run: officecli close \"{file}\""
                };
            }

            OfficeCli.BlankDocCreator.Create(file);
            var fullCreatedPath = Path.GetFullPath(file);

            // Best-effort: auto-start a short-lived resident process so
            // follow-up commands on this freshly-created file hit the
            // in-memory handler instead of re-opening from disk each time.
            // Uses a 60s idle timeout (much shorter than `open`'s default
            // 12min) so a stray `create` with no follow-up exits quickly.
            // Failure here does NOT fail the command — the file is already
            // on disk and all other commands still work via direct open.
            var residentStarted = TryStartResidentProcess(fullCreatedPath, idleSeconds: 60, out var residentErr);
            var residentSuffix = residentStarted
                ? " (resident started, auto-close in 60s idle)"
                : "";

            if (json)
            {
                Console.WriteLine(OutputFormatter.WrapEnvelopeText($"Created: {fullCreatedPath}{residentSuffix}"));
            }
            else
            {
                Console.WriteLine($"Created: {file}{residentSuffix}");
                if (!residentStarted && !string.IsNullOrEmpty(residentErr))
                {
                    Console.Error.WriteLine($"Note: resident auto-start failed ({residentErr}); falling back to direct file access.");
                }
                if (Path.GetExtension(file).Equals(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"  totalSlides: 0");
                    Console.WriteLine($"  slideWidth: {Core.EmuConverter.FormatEmu(12192000)}");
                    Console.WriteLine($"  slideHeight: {Core.EmuConverter.FormatEmu(6858000)}");
                }
            }
            return 0;
        }, json); });

        return createCommand;
    }

    private static Command BuildMergeCommand(Option<bool> jsonOption)
    {
        var mergeTemplateArg = new Argument<string>("template") { Description = "Template file path (.docx, .xlsx, .pptx) with {{key}} placeholders" };
        var mergeOutputArg = new Argument<string>("output") { Description = "Output file path" };
        var mergeDataOpt = new Option<string>("--data") { Description = "JSON data or path to .json file", Required = true };
        var mergeCommand = new Command("merge", "Merge template with JSON data, replacing {{key}} placeholders");
        mergeCommand.Add(mergeTemplateArg);
        mergeCommand.Add(mergeOutputArg);
        mergeCommand.Add(mergeDataOpt);
        mergeCommand.Add(jsonOption);

        mergeCommand.SetAction(result => { var json = result.GetValue(jsonOption); return SafeRun(() =>
        {
            var template = result.GetValue(mergeTemplateArg)!;
            var output = result.GetValue(mergeOutputArg)!;
            var dataArg = result.GetValue(mergeDataOpt)!;

            var data = Core.TemplateMerger.ParseMergeData(dataArg);
            var mergeResult = Core.TemplateMerger.Merge(template, output, data);

            if (json)
            {
                var jsonObj = new System.Text.Json.Nodes.JsonObject
                {
                    ["success"] = true,
                    ["output"] = Path.GetFullPath(output),
                    ["replacedKeys"] = mergeResult.UsedKeys.Count,
                    ["unresolvedPlaceholders"] = new System.Text.Json.Nodes.JsonArray(
                        mergeResult.UnresolvedPlaceholders.Select(p => (System.Text.Json.Nodes.JsonNode)p).ToArray())
                };
                Console.WriteLine(jsonObj.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = false }));
            }
            else
            {
                Console.WriteLine($"Merged: {output}");
                Console.WriteLine($"  Replaced keys: {mergeResult.UsedKeys.Count}");
                if (mergeResult.UnresolvedPlaceholders.Count > 0)
                {
                    Console.Error.WriteLine($"  Warning: {mergeResult.UnresolvedPlaceholders.Count} unresolved placeholder(s):");
                    foreach (var p in mergeResult.UnresolvedPlaceholders)
                        Console.Error.WriteLine($"    - {{{{{p}}}}}");
                }
            }
            return 0;
        }, json); });

        return mergeCommand;
    }
}
