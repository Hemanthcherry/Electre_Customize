using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Threading;
using Electre_Customize_DotNet.Logs;
using Electre_Customize_DotNet.MainOperation;

namespace Electre_Customize_DotNet.Helpers.ExcelHelpers
{
    /// <summary>
    /// Runs the per-loom / per-sheet report generators either exactly as before (one Excel instance, one item after
    /// the other, on the calling thread) or - when App.config ExcelWorkerCount is greater than 1 - on several
    /// separate Excel processes at once.
    ///
    /// Every loom / sheet report is an independent workbook, so this is the only way to use more than one CPU core for
    /// the Excel part of a huge job (Excel itself is single threaded). Each worker thread is STA, owns its own
    /// modExcel with its own Excel process, and never touches the caller's instance; the shared data the report
    /// code reads (arrFT_CwithBCProject, ElecCollection ...) is read-only during report generation.
    /// A worker whose Excel start-up returns an instance that is already in use is dropped rather than shared,
    /// and anything the workers could not take is processed on the caller's thread with the caller's instance.
    /// </summary>
    internal static class ExcelParallel
    {
        public const int MaxWorkers = 8;

        public static int ConfiguredWorkers()
        {
            if (int.TryParse(ConfigurationManager.AppSettings["ExcelWorkerCount"], out int n) && n > 1)
            {
                return Math.Min(n, MaxWorkers);
            }
            return 1;
        }

        public static void ForEach<TItem>(IReadOnlyList<TItem> items, modExcel sharedExcel, Action<modExcel, TItem> body)
        {
            ForEach<TItem, object>(items, sharedExcel, () => null, (excel, item, _) => body(excel, item));
        }

        /// <param name="stateFactory">Creates the per-worker state (for example a reusable row buffer). With one worker it is created once, like the old locals.</param>
        public static void ForEach<TItem, TState>(
            IReadOnlyList<TItem> items,
            modExcel sharedExcel,
            Func<TState> stateFactory,
            Action<modExcel, TItem, TState> body)
        {
            int workers = Math.Min(ConfiguredWorkers(), items.Count);

            if (workers <= 1)
            {
                // the original behaviour: same thread, same Excel instance, same order
                var state = stateFactory();
                foreach (var item in items)
                {
                    body(sharedExcel, item, state);
                }
                return;
            }

            // read on the caller's thread: the caller's Excel object must not be touched from the workers
            var usedWindows = new ConcurrentDictionary<long, bool>();
            long callerWindow = WindowOf(sharedExcel);
            if (callerWindow != 0)
            {
                usedWindows[callerWindow] = true;
            }

            var next = new NextIndex();
            var threads = new List<Thread>(workers);
            Logging.Info($"Excel workers: {workers} parallel Excel instances for {items.Count} items");

            for (int w = 0; w < workers; w++)
            {
                var thread = new Thread(() => RunWorker(items, stateFactory, body, usedWindows, next))
                {
                    Name = "ExcelWorker" + (w + 1),
                };
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                threads.Add(thread);
            }

            foreach (var thread in threads)
            {
                thread.Join();
            }

            // items no worker took (e.g. no worker could start its own Excel): finish them the old way
            var remainingState = stateFactory();
            while (true)
            {
                int idx = Interlocked.Increment(ref next.Value) - 1;
                if (idx >= items.Count)
                {
                    break;
                }
                body(sharedExcel, items[idx], remainingState);
            }
        }

        private static void RunWorker<TItem, TState>(
            IReadOnlyList<TItem> items,
            Func<TState> stateFactory,
            Action<modExcel, TItem, TState> body,
            ConcurrentDictionary<long, bool> usedWindows,
            NextIndex next)
        {
            modExcel excel = null;
            bool ownsInstance = false;
            try
            {
                excel = new modExcel();
                excel.TemplateReadOnly = true;   // several Excel processes open the same template at once
                excel.InitiateExcel();
                if (excel.ExcelApp == null)
                {
                    return;
                }

                long window = WindowOf(excel);
                if (window != 0 && !usedWindows.TryAdd(window, true))
                {
                    Logging.Warning("Excel worker skipped: Excel returned an instance that is already in use.");
                    return;
                }
                ownsInstance = true;

                var state = stateFactory();
                while (true)
                {
                    int idx = Interlocked.Increment(ref next.Value) - 1;
                    if (idx >= items.Count)
                    {
                        break;
                    }

                    try
                    {
                        body(excel, items[idx], state);
                    }
                    catch (Exception ex)
                    {
                        Logging.Error($"Excel worker item {idx}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logging.Error("Excel worker failed: " + ex.Message);
            }
            finally
            {
                if (excel?.ExcelApp != null)
                {
                    try
                    {
                        if (ownsInstance)
                        {
                            excel.releaseExcel();   // Quit + release
                        }
                        else
                        {
                            Marshal.ReleaseComObject(excel.ExcelApp);   // never Quit an instance that is not ours
                        }
                    }
                    catch
                    {
                    }
                }
            }
        }

        // shared work counter (a ref to a captured local cannot be used inside the worker lambda)
        private sealed class NextIndex
        {
            public int Value;
        }

        private static long WindowOf(modExcel excel)
        {
            try
            {
                return excel?.ExcelApp != null ? (long)excel.ExcelApp.Hwnd : 0;
            }
            catch
            {
                return 0;
            }
        }
    }
}
