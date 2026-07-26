// ESAPI v18 (writeable) – Breast Optimisation Structure Generator (GUI)
//
// Original author: Joshua Southwell, Medical Physicist (Australian Volunteer at LMH)
// Modifications by: Ratha SAN, 2026-04
// Refinements by: Gemini (2026-04), reviewed/cleaned 2026-06
//
// DESCRIPTION:
//   Specialised version of the generic optimisation script tailored for breast/chest wall.
//   Incorporates Virtual Bolus logic using strict asymmetric margin expansions.
//
// CHANGELOG:
//   v3.0.0.18 – Suffix eradication for Ovl; z-prefix naming cleanup.
//   v3.0.0.19 – 3-tier avoidance architecture; Avoid checkbox is target-specific override only.
//   v3.0.0.20 – Removed global zAvoidance_sum; dose-level gradient shells only.
//   v3.0.0.24 – Virtual Bolus pipeline redesigned (5-step geometry):
//               z_Virtual_PTV (ant+lat expand = input+2mm), z_Virtual_Bolus_raw (outside body),
//               Body_new (body Or Virtual_PTV, updates External in-place), z_Virtual_Bolus
//               (raw bolus ∩ Body_new−2mm), z_Virtual_PTV_Opt (bolus expanded post+inf 5mm,
//               capped to Body_new−2mm).
//   v3.0.0.30 – Physical Bolus support:
//               New "Physical Bolus" checkbox in laterality panel.
//               When ticked, script finds existing BOLUS-type structure and uses
//               Body_with_Bolus = body Or Bolus as the skin baseline for Step 5.
//               Virtual bolus is then generated on top of the physical bolus surface.
//               Aborts with error if no BOLUS structure found in structure set.
//   v3.0.0.31 – Physical Bolus optimisation structures:
//               When Physical Bolus is ticked, Step 5 now also creates:
//                 Bolus_physical  = copy of the found BOLUS structure (type CONTROL -
//                                   Eclipse's scripting API refuses to AddStructure
//                                   with dicomType "BOLUS"; that type can only be
//                                   created via Insert > New Bolus... in the UI).
//                 Bolus_phys_Opt  = overlap (intersection) of Bolus_physical and
//                                   z_Virtual_PTV (type CONTROL, matching
//                                   Bolus_physical so both group/sort together in
//                                   the Structures list). Named "Bolus_phys_Opt"
//                                   rather than "Bolus_physical_Opt" to stay within
//                                   Eclipse's 16-character structure ID limit.
//               Body_new is unchanged (already = Body_with_Bolus Or z_Virtual_PTV,
//               i.e. Body + Bolus + z_Virtual_PTV when Physical Bolus is present).
//   v3.0.0.32 – Physical Bolus thickness dropdown + simplified formulas:
//               A physical bolus thickness dropdown (5/10/15/20mm) is enabled
//               alongside "Physical Bolus present". When ticked, Step 1's
//               per-target expansion (totalMm) becomes bolusMm input + selected
//               physical thickness (no extra +2mm buffer), replacing the
//               non-physical-bolus totalMm = bolusMm + VB_EXTRA_EXPAND_MM.
//               Step 4/5 also use simplified, physical-bolus-specific formulas:
//                 z_Virtual_Bolus   = z_Virtual_PTV Sub Bolus_physical Sub Body.
//                 z_Virtual_PTV_Opt = union of the z_PTV_opt targets feeding
//                                     z_Virtual_PTV, expanded ant+lat by
//                                     VB_PHYS_OPT_EXPAND_MM (4mm), capped to
//                                     Body_new.
//               The original 5-step rawBolus/skin-crop pipeline (Body_with_Bolus,
//               VB_SKIN_CROP_MM, VB_OPT_SKIN_CROP_MM, VB_OPT_INWARD_MM) still
//               runs unchanged when Physical Bolus is NOT ticked.
//   v3.1.0.0  – Added "RCC" tab (independent, additive – the original single-tab
//               UI/workflow described above is unchanged and lives in the
//               "Breast Optimiser" tab). RCC implements the "RCC Optimization
//               Cropping Method" clinical cheat sheet formulas as a second,
//               self-contained logic engine operating on its own target/OAR
//               selections inside the same dialog:
//                 - zPTV Opti      : %Diff = (Rx-OARMax)/Rx*100, Crop = %Diff/ZoneA
//                                    (strict Max-Dose OAR sparing crop).
//                 - SIB shave      : %Diff = (RxHigh-RxLow)/RxHigh*100, Crop = %Diff/ZoneB
//                                    (zPTV{Low}_Opti = PTVLow Sub PTVHigh+Crop).
//                 - z-ring_sib1/2  : rings at the 85%/65% isodose of the lowest
//                                    ticked Rx, built per-level at ZoneB/ZoneC
//                                    falloff (ring1 floor 3.0mm), non-overlapping
//                                    shells (ring2 sits outside ring1).
//                 - zOAR-in-PTV1/2 : 2mm/4mm nested shells of OAR-in-Prv overlap
//                                    for iterative nested-ring sparing, plus a
//                                    fully-subtracted zPTV Opti per target/OAR pair.
//               See RCC constants/classes/methods below (all new; nothing in the
//               original Script.Execute / StructureProcessor / OptimisationStructureWindow
//               original-tab code paths was modified).
//   v4.0.0.0  – Split into three tabs, each backed by its own independent
//               target/OAR selections (SiteTabController.TabKind), so ticks in
//               one tab never leak into another:
//                 - "Generic"     : the shared Step1-9 pipeline (StructureProcessor,
//                                   unchanged) with laterality + virtual/physical
//                                   bolus controls removed - all-site PTV/OAR
//                                   crop automation.
//                 - "Breast Opto" : the original breast/chest-wall workflow
//                                   (laterality, Virtual Bolus, Physical Bolus),
//                                   unchanged, just relocated from the old
//                                   single "Generic"-labelled tab onto its own
//                                   "Breast Opto" tab.
//                 - "RCC"         : the RCC Optimization Cropping Method engine,
//                                   rebuilt onto the same grid-based UI shell as
//                                   the Generic tab (top bar / targets grid /
//                                   OAR grid / bottom bar) instead of its previous
//                                   bespoke layout, with Rx/Max-Dose/Nested-sparing
//                                   now ticked directly on TargetDoseRow/OrganRow
//                                   (RccTargetRow/RccOarRow removed) plus the
//                                   crop-distance matrix and advanced SIB/ring/
//                                   nested plan preview retained as additional
//                                   panels within that shell.
//   v4.1.0.0  – RCC engine rework + dead-code/duplication cleanup across all
//               three tabs:
//                 - RCC's "Create SIB / Ring / Nested Structures" button now
//                   runs a six-step Eval -> Opt -> Opt Sum -> Ring1 -> Ring2 ->
//                   SIB shave -> Rind pipeline (mirrors Step1_EvalPtv /
//                   Step2_OptPtv / Step3b_GlobalOptPtvSum / Step9_Rings, driven
//                   by RCC's falloff-zone formulas instead of fixed margins).
//                   Ring1/Ring2 now build from a shared PTV_Opt_Sum instead of
//                   each raw ticked PTV; SIB shave now crops PTV_Opt (not the
//                   raw PTV); new z{target}_Rind = outer 5mm shell of each
//                   target's final (post-shave) PTV_Opt. §7 nested OAR-in-PTV
//                   sparing and the OAR-max-dose Auto-Crop pipeline (§2) are
//                   unchanged.
//                 - RCC "Auto-Crop" renamed "Generate Structure"; validates
//                   required input up front (alerts exactly what's missing)
//                   and closes the dialog on a clean run instead of leaving it
//                   open.
//                 - RCC OAR grid: removed the separate "Crop" tick - an OAR
//                   now participates the moment a Max Dose is entered - and
//                   added mutually-exclusive Small/Large tick columns so each
//                   OAR can pick its own Zone A falloff rate (10%/mm small,
//                   5%/mm large, both configurable).
//                 - RCC layout: falloff-zone rates moved into their own
//                   stacked "FALLOFF ZONE" card below the Targets grid; the
//                   Crop Distance Matrix moved there too (below the falloff
//                   card); the Organs grid now fills the whole right panel
//                   (shares a layout cell with the Advanced plan preview
//                   instead of being height-capped alongside the matrix).
//                 - Removed dead code found across the three-tab split:
//                   StructureProcessor's unused `externals` field/param,
//                   SiteTabController's write-only `_inAdvMode` field and
//                   unused `Vm` property, and several DataGridColumn fields
//                   that were assigned once and never read again. Extracted
//                   the repeated "header-tick -> foreach-set -> refresh"
//                   column-building blocks (Ovl/Opt/PRV/Avoid/Crop?/etc.) into
//                   one shared AddBoolColumn<T>() helper.
//   v4.2.0.0  – RCC: single-target and multi-target (SIB) unified into one
//               process/button, and Falloff Zone laid out as one row:
//                 - Single ticked target and multi-target (SIB) selections now
//                   go through the exact same "Generate Structure" click
//                   (DoRccGenerateStructure) instead of two separate buttons/
//                   modes: Eval -> Opt -> §2 OAR max-dose crop -> Opt Sum ->
//                   [2+ targets only: Ring1 -> Ring2 -> SIB shave] -> Rind ->
//                   [optional] §7 nested sparing, always in that order. A
//                   single target still gets Eval/Opt/OAR-crop/Rind; it just
//                   has no "other" PTV to build a ring or SIB-shave against,
//                   so only those two steps are skipped.
//                 - §2 OAR max-dose crop now applies directly to each target's
//                   PTV_Opt (in place, same z{target}_Opt id) instead of the
//                   raw PTV, and instead of being its own separate pipeline -
//                   required for both the single-target and multi-target
//                   cases alike, and feeds into Opt_Sum/Ring/SIB/Rind so they
//                   all reflect the OAR-sparing already applied.
//                 - SIB shave now updates z{target}_Opt in place too (cascades
//                   correctly down the dose ladder) instead of creating a
//                   separate "z{Low}_Opti" structure - there is one evolving
//                   PTV_Opt per target now, not two similarly-named structures.
//                 - Removed the "Create SIB / Ring / Nested Structures" button
//                   and the Auto-Crop/Advanced mode split; RCC's View combo is
//                   now a pure OARs/Advanced-Plan-Preview toggle with no effect
//                   on which button is visible, since there's only one.
//                 - FALLOFF ZONE card laid out as a single row with one column
//                   per zone (A-small, A-large, B, C) instead of stacked rows.
//   v4.3.0.0  – RCC: PTV_Opt is now the only optimisation target structure,
//               smoothed once at the end, forced to high resolution:
//                 - Removed the "_Opti" naming/structure entirely. Section 7's
//                   nested OAR-in-PTV sparing now crops the SAME z{target}_Opt
//                   used by every other step (Eval->Opt->§2 OAR crop->SIB
//                   shave->§7 nested crop), matching how §2/SIB already mutate
//                   it in place. There is exactly one PTV_Opt per target for
//                   the whole RCC engine, built from PTV_Eval, never a second
//                   "_Opti" result.
//                 - Smoothing (expand/contract by SMOOTH_MM, level 3) moved
//                   from PTV_Opt's initial creation to a single pass over every
//                   target's FINAL PTV_Opt, right after §7 nested crop and
//                   right before Rind is built - smoothing at creation was
//                   being silently undone by every later crop (§2, SIB, §7).
//                 - Added EnsureRccHighRes() and called it on every structure
//                   the RCC pipeline creates or reuses (PTV_Eval, PTV_Opt,
//                   PTV_Opt_Sum, Ring1, Ring2, Rind, zOAR-in-PTV1/2) so all
//                   RCC output is high resolution.
//   v4.4.0.0  – RCC: PTV_Opt no longer expands past PTV_Eval; rings renamed;
//               §7 nested sparing generalised to N variable-thickness shells:
//                 - Step 2 no longer expands PTV_Opt +2mm past PTV_Eval - Opt
//                   now starts out identical to Eval (capped to Body) and is
//                   reshaped from there by every later crop (§2 OAR max-dose,
//                   SIB shave, §7 nested). EVAL_TO_OPT_EXPAND_MM stays in use
//                   by the separate Generic/Breast Opto pipeline only.
//                 - Renamed the Zone B/C isodose rings from "z-ring_sib1"/
//                   "z-ring_sib2" to "z_Ring_1"/"z_Ring_2".
//                 - §7 nested OAR-in-PTV sparing no longer hardcodes 2
//                   shells at a fixed 2mm/4mm step. Each OAR ticked "Nested
//                   §7" now has its own "Nested Thickness (mm)" input (next
//                   to the tick column); BuildRccNestedShells steps outward
//                   from that OAR's surface by that thickness, building one
//                   z_oar_in_ptv_hr{level} structure per level (e.g.
//                   z_parotidL_in_ptv_hr1, _hr2, ...), unioned across every
//                   ticked target the OAR overlaps, until a level no longer
//                   overlaps any of them - so shell count follows the actual
//                   OAR/target geometry instead of a fixed count of two.
//   v4.5.0.0  – RCC: removed the level-3 expand/contract smoothing pass added
//               in v4.3.0.0. It ran on both PTV_Opt_Sum (in BuildRccOptSum)
//               and every target's final PTV_Opt (end of
//               DoRccGenerateStructure); either one nudges the boundary past
//               what the eval+crop formula computes. PTV_Opt is now exactly
//               PTV_Eval cropped by §2/SIB/§7, nothing else - Rind still
//               builds from that same (now unsmoothed) PTV_Opt.
//               SMOOTH_OPT_TARGET/SMOOTH_MM/SmoothStructureByExpandContract
//               remain in use by the separate Generic/Breast Opto pipeline.
//   v4.6.0.0  – RCC: "Generate Structure" now closes the dialog immediately
//               after it runs, regardless of outcome. Previously a creation
//               error kept the dialog open (with a summary shown) so it
//               could be retried in place; it now shows that same summary
//               first, then closes either way. Missing/invalid input still
//               blocks the run before anything is generated and keeps the
//               dialog open to fix it - only the post-generation branch
//               changed.
//   v4.7.0.0  – RCC: fixed z{target}_Rind to actually be the outer shell its
//               own name/comments always claimed. BuildRccRind previously
//               contracted PTV_Opt inward by RCC_RIND_INWARD_MM and
//               subtracted that from PTV_Opt itself, producing an INNER
//               5mm shell just inside the Opt boundary despite being
//               documented as "outer". It now expands PTV_Opt outward by
//               RCC_RIND_OUTWARD_MM (renamed from RCC_RIND_INWARD_MM) and
//               subtracts the original PTV_Opt from that, so Rind is the
//               solid 5mm shell surrounding PTV_Opt, capped to Body.
//   v4.8.0.0  – RCC: corrected z{target}_Rind again, per explicit direction
//               ("rind = opt - 5mm margin"). v4.7.0.0's outer-shell boolean
//               (PTV_Opt expanded 5mm, minus PTV_Opt) is replaced by a plain
//               negative margin: Rind = PTV_Opt contracted inward by
//               RCC_RIND_MARGIN_MM (renamed from RCC_RIND_OUTWARD_MM),
//               capped to Body - a solid, smaller volume, not a boolean
//               shell/ring at all.
//   v4.9.0.0  – RCC: ApplyRccMaxDoseCropToOpt no longer silently skips the
//               §2 OAR max-dose crop when it fails. Two paths used to
//               `continue` with no message: PTV_Opt missing from
//               optByTarget, and the expanded-OAR union coming back null
//               (e.g. every expanded-OAR temp ended up empty). Both now add
//               to `errors` so a failed crop shows up in the Generate
//               Structure summary instead of the dialog just closing as if
//               nothing was wrong. Note this only fires on generation
//               failures - if an OAR's Max Dose is >= its target's Rx, the
//               formula correctly computes zero crop pairs for that pair
//               (check the Crop Distance Matrix), which is not an error.
//
// KNOWN LIMITATIONS (not yet fixed in this version):
//   - _zOptDoseSum is keyed by dose (double) only. If two groups share the same dose level
//     but differ by suffix, Step3a will overwrite the first entry and Step9 rings will only
//     see one structure per dose. To fix properly: change _zOptDoseSum to Dictionary<OptKey, Structure>.
//   - SafeAsymmetricMargin uses y1=anterior expansion. Verify against your linac IEC patient
//     coordinate convention before clinical use.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

[assembly: AssemblyVersion("4.9.0.0")]
[assembly: AssemblyFileVersion("4.9.0.0")]
[assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        // =======================
        // CONFIG
        // =======================
        private const double BODY_CONTRACT_MM = 3.0;
        private const double EVAL_TO_OPT_EXPAND_MM = 2.0;
        private const double LOWER_SUBTRACT_EXTRA_MM = 1.0;
        private const double RING_OUTER_EXPAND_MM = 10.0;
        private const double DEFAULT_PRV_MARGIN_MM = 2.0;
        private const double AVOIDANCE_MARGIN_MM = 35.0;

        // Virtual Bolus geometry constants
        private const double VB_EXTRA_EXPAND_MM = 2.0;  // added to user bolus input for Virtual_PTV expansion (no physical bolus)
        private const double VB_SKIN_CROP_MM = 2.0;  // trim outer shell from Body_new for z_Virtual_Bolus
        private const double VB_OPT_SKIN_CROP_MM = 4.0;  // trim outer shell from Body_new for z_Virtual_PTV_Opt (no physical bolus)
        private const double VB_OPT_INWARD_MM = 5.0;  // expansion magnitude for Virtual_PTV_Opt (no physical bolus)
        private const double VB_PHYS_OPT_EXPAND_MM = 4.0;  // ant+lat expansion of z_PTV_opt for z_Virtual_PTV_Opt (physical bolus mode)

        private static readonly double[] PhysicalBolusThicknessOptionsMm = { 5.0, 10.0, 15.0, 20.0 };

        // ---------------------------------------------------------------
        // RCC LOGIC ENGINE CONFIG
        // "PTV Cropping & Optimization Formulas" clinical cheat sheet, v1.0
        // ---------------------------------------------------------------
        // Zone A splits by OAR size per the cheat sheet's own note ("Falloff:
        // 10%/mm small volumes - 5%/mm large"): a small/critical OAR gets the
        // steeper (higher %/mm) rate, a large OAR the shallower one (which
        // yields a bigger crop for the same %Diff). Selected per-row via
        // OrganRow.IsSmallOrgan/IsLargeOrgan.
        private const double RCC_ZONE_A_SMALL_DEFAULT_PCT_PER_MM = 10.0; // small/critical OAR max-dose crop falloff
        private const double RCC_ZONE_A_LARGE_DEFAULT_PCT_PER_MM = 5.0;  // large OAR max-dose crop falloff
        private const double RCC_ZONE_B_DEFAULT_PCT_PER_MM = 5.0;  // SIB shave / z_Ring_1 falloff
        private const double RCC_ZONE_C_DEFAULT_PCT_PER_MM = 2.7;  // z_Ring_2 (far-target) falloff
        private const double RCC_RING1_ISO_FRACTION = 0.85;        // z_Ring_1 target = 85% x lowest ticked Rx
        private const double RCC_RING2_ISO_FRACTION = 0.65;        // z_Ring_2 target = 65% x lowest ticked Rx
        private const double RCC_RING1_MIN_CROP_MM = 3.0;          // z_Ring_1 minimum crop distance
        private const double RCC_NESTED_RING_STEP_MM = 2.0;        // z_oar_in_ptv_hr# default shell thickness (used when OrganRow.NestedThicknessMm is blank/invalid)
        private const int RCC_NESTED_MAX_LEVELS = 30;               // safety cap on how many z_oar_in_ptv_hr# shells one OAR can produce
        private const double RCC_RIND_MARGIN_MM = 5.0;              // z{target}_Rind = PTV_Opt contracted inward by this margin

        private const bool SMOOTH_OPT_TARGET = true;
        private const double SMOOTH_MM = 3.0;
        private const bool OVERLAP_IS_INTERSECTION = true;

        private const string ID_OPT_TV_SUM = "z_PTV_opt_sum";

        // Varian Eclipse hard limit for structures in a single StructureSet
        private const int MAX_STRUCTURES = 255;

        // Priority OARs for breast – used for auto-tick and sort in UI
        private static readonly string[] PRIORITY_OARS =
        {
            "brachial", "esophagus", "esoph", "lung", "thyroid", "trachea"
        };

        // ==================================================================
        // MEMORY HELPER
        // ==================================================================
        private sealed class TempGuard : IDisposable
        {
            private readonly StructureSet _ss;
            private readonly List<Structure> _temps = new List<Structure>();

            public TempGuard(StructureSet ss) { _ss = ss; }

            public Structure Add(Structure s) { if (s != null) _temps.Add(s); return s; }

            public void Dispose()
            {
                foreach (var t in _temps)
                {
                    try
                    {
                        var s = _ss.Structures.FirstOrDefault(x => x.Id == t.Id);
                        if (s != null) _ss.RemoveStructure(s);
                    }
                    catch { /* best-effort cleanup */ }
                }
                _temps.Clear();
            }
        }

        // ==================================================================
        // ENTRY POINT
        // ==================================================================
        public void Execute(ScriptContext context)
        {
            if (context?.Patient == null || context.StructureSet == null)
            {
                MessageBox.Show("No patient/structure set loaded.");
                return;
            }

            var ss = context.StructureSet;
            if (ss.Image == null)
            {
                MessageBox.Show("No primary image found for this StructureSet.");
                return;
            }

            var allNonEmpty = ss.Structures.Where(s => s != null && !s.IsEmpty).ToList();

            var externals = allNonEmpty
                .Where(s => string.Equals(s.DicomType, "EXTERNAL", StringComparison.OrdinalIgnoreCase)
                         || ContainsToken(s.Id, "EXTERNAL")
                         || ContainsToken(s.Id, "BODY"))
                .Distinct()
                .ToList();

            var targetCandidates = allNonEmpty
                .Where(IsTargetByName)
                .OrderBy(s => s.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (targetCandidates.Count == 0)
            {
                MessageBox.Show("No target volumes found by name (PTV/CTV/GTV/LN).");
                return;
            }

            var oarCandidates = allNonEmpty
                .Where(s => !IsTargetByName(s))
                .Where(s => !string.Equals(s.DicomType, "EXTERNAL", StringComparison.OrdinalIgnoreCase))
                .Where(s => !ContainsToken(s.Id, "EXTERNAL") && !ContainsToken(s.Id, "BODY"))
                .OrderBy(s => s.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();

            context.Patient.BeginModifications();

            var vmGeneric = new UiModel(targetCandidates, oarCandidates, externals);
            var vmBreast = new UiModel(targetCandidates, oarCandidates, externals);
            var vmRcc = new UiModel(targetCandidates, oarCandidates, externals);
            var win = new OptimisationStructureWindow(vmGeneric, vmBreast, vmRcc, ss);

            if (win.ShowDialog() != true || win.ConfirmedVm == null) return;

            var processor = new StructureProcessor(ss, win.ConfirmedVm, targetCandidates);
            processor.Run();
        }

        // ==================================================================
        // VALUE TYPES
        // ==================================================================
        private struct OptKey : IEquatable<OptKey>
        {
            public double DoseGy { get; }
            public string Suffix { get; }

            public OptKey(double doseGy, string suffix)
            {
                DoseGy = doseGy;
                Suffix = (suffix ?? "").Trim();
            }

            public bool Equals(OptKey other) =>
                DoseGy == other.DoseGy &&
                string.Equals(Suffix, other.Suffix, StringComparison.OrdinalIgnoreCase);

            public override bool Equals(object obj) => obj is OptKey ok && Equals(ok);

            public override int GetHashCode()
            {
                unchecked { return (DoseGy.GetHashCode() * 397) ^ Suffix.ToUpperInvariant().GetHashCode(); }
            }
        }

        private class TargetDosePair
        {
            public string TargetId { get; set; }
            public double DoseGy { get; set; }
            public string Suffix { get; set; }
            public double? BolusMm { get; set; }
            public bool CreateAvoidance { get; set; }
        }

        // ==================================================================
        // STRUCTURE GENERATION PROCESSOR
        // ==================================================================
        private class StructureProcessor
        {
            private readonly StructureSet _ss;
            private readonly UiModel _vm;
            private readonly List<Structure> _targetCandidates;

            private readonly StringBuilder _progress = new StringBuilder();
            private int _createdCount;
            private SliceRecontourFallback _fb;

            // Keyed by OptKey (dose + suffix)
            private readonly Dictionary<OptKey, Structure> _zEval = new Dictionary<OptKey, Structure>();
            private readonly Dictionary<OptKey, Structure> _zOpt = new Dictionary<OptKey, Structure>();
            // Keyed by dose only – see KNOWN LIMITATIONS note at top of file
            private readonly Dictionary<double, Structure> _zOptDoseSum = new Dictionary<double, Structure>();

            public StructureProcessor(
                StructureSet ss, UiModel vm, List<Structure> targetCandidates)
            {
                _ss = ss;
                _vm = vm;
                _targetCandidates = targetCandidates;
            }

            private void LogCreated(string id)
            {
                _progress.AppendLine($"  OK: {id}");
                _createdCount++;
            }

            private void LogSection(string name)
            {
                _progress.AppendLine();
                _progress.AppendLine($"-- {name} --");
            }

            public void Run()
            {
                var selectedExternal = _vm.SelectedExternal ?? FindExternalFallback(_ss);
                if (selectedExternal == null || selectedExternal.IsEmpty)
                {
                    MessageBox.Show("No External/Body structure selected or found.");
                    return;
                }

                // IsLeftSided is only ever set via the Breast Opto tab's laterality
                // radio buttons; the Generic tab never shows them and always sends
                // an empty bolusRequests list, so isLeft is never dereferenced by
                // Step5_VirtualBolus in that case - default to false rather than
                // crashing on Nullable<bool>.Value.
                bool isLeft = _vm.IsLeftSided ?? false;
                var targetDoseRows = _vm.TargetDoseRows.Where(r => r.IsSelected).ToList();

                var targetById = _targetCandidates
                    .Where(t => targetDoseRows.Any(r =>
                        string.Equals(r.TargetId, t.Id, StringComparison.OrdinalIgnoreCase)))
                    .ToDictionary(s => s.Id, s => s, StringComparer.OrdinalIgnoreCase);

                var selectedTargets = targetById.Values.ToList();
                var organRows = _vm.OrganRows.ToList();

                var targetDosePairs = targetDoseRows
                    .Select(r => new TargetDosePair
                    {
                        TargetId = r.TargetId,
                        DoseGy = r.ParsedDoseGy.Value,
                        Suffix = (r.Suffix ?? "").Trim(),
                        BolusMm = r.ParsedBolusMm,
                        CreateAvoidance = r.CreateAvoidance
                    })
                    .ToList();

                var groupKeys = targetDosePairs
                    .Select(x => new OptKey(x.DoseGy, x.Suffix))
                    .Distinct()
                    .OrderByDescending(k => k.DoseGy)
                    .ThenBy(k => k.Suffix, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var targetsByGroup = new Dictionary<OptKey, List<Structure>>();
                foreach (var k in groupKeys)
                {
                    var ids = targetDosePairs
                        .Where(x => x.DoseGy == k.DoseGy &&
                                    string.Equals(x.Suffix, k.Suffix, StringComparison.OrdinalIgnoreCase))
                        .Select(x => x.TargetId)
                        .ToList();
                    targetsByGroup[k] = ids
                        .Select(id => targetById.TryGetValue(id, out var s) ? s : null)
                        .Where(s => s != null && !s.IsEmpty)
                        .ToList();
                }

                var doseLevels = groupKeys
                    .Select(k => k.DoseGy)
                    .Distinct()
                    .OrderByDescending(x => x)
                    .ToList();

                var bolusRequests = targetDosePairs
                    .Where(r => r.BolusMm.GetValueOrDefault() > 0)
                    .ToList();

                // -------------------------------------------------------
                // Resolve physical bolus structure (v3.0.0.30)
                // -------------------------------------------------------
                Structure physicalBolus = null;
                if (_vm.HasPhysicalBolus)
                {
                    physicalBolus = _ss.Structures.FirstOrDefault(s =>
                        s != null && !s.IsEmpty &&
                        string.Equals(s.DicomType, "BOLUS", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(s.Id, "Bolus_physical", StringComparison.OrdinalIgnoreCase));

                    // CommitSelections() already validated this exists; guard here defensively.
                    if (physicalBolus == null)
                    {
                        MessageBox.Show(
                            "Physical Bolus was selected but no BOLUS structure was found in the structure set.\n\nScript aborted.",
                            "Missing BOLUS Structure", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                    _progress.AppendLine($"  Physical Bolus: {physicalBolus.Id} (DICOM type BOLUS)");
                }

                try
                {
                    _fb = new SliceRecontourFallback();

                    using (var globalTg = new TempGuard(_ss))
                    {
                        var extMinus3 = SafeMargin(selectedExternal.SegmentVolume, -BODY_CONTRACT_MM);
                        var bodyMinus3 = SafeBoolean(_ss, extMinus3, selectedExternal.SegmentVolume,
                                            BoolOp.And, selectedExternal, selectedExternal,
                                            null, _fb, "BodyMinus3", globalTg);

                        Step1_EvalPtv(groupKeys, targetsByGroup, selectedExternal, bodyMinus3);
                        Step2_OptPtv(groupKeys, selectedExternal);
                        Step3a_DoseOptPtvSum(groupKeys, doseLevels, selectedExternal);
                        Step3b_GlobalOptPtvSum(selectedTargets, selectedExternal);
                        Step4_Avoidance(groupKeys, doseLevels, targetDosePairs, selectedExternal);

                        // Step5: pass physicalBolus (null if not used).
                        // When non-null, Step5 builds Body_with_Bolus = body Or physicalBolus
                        // and uses it as the skin baseline. Steps 6-9 still use selectedExternal.
                        Step5_VirtualBolus(bolusRequests, selectedExternal, isLeft, globalTg,
                            physicalBolus, _vm.PhysicalBolusThicknessMm);

                        Step6_Overlaps(organRows, doseLevels, selectedExternal);
                        Step7_OptOars(organRows, groupKeys, selectedExternal);
                        Step8_Prvs(organRows, selectedExternal);
                        Step9_Rings(doseLevels, selectedExternal, extMinus3);
                    }

                    MessageBox.Show(
                        $"Script complete.\n\nTotal structures created: {_createdCount}\n\nDetails:\n{_progress}",
                        "Script Progress", MessageBoxButton.OK, MessageBoxImage.Information);

                    var recontoured = _fb.CreatedStructureIdsWithFallback.ToList();
                    if (recontoured.Count > 0)
                    {
                        MessageBox.Show(
                            "WARNING: One or more boolean operations required robust slice-by-slice geometry fallbacks.\n" +
                            "This happens safely in the background, but review these structures:\n\n" +
                            string.Join("\n", recontoured),
                            "Geometry Fallback Triggered", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Script FAILED.\n\nError: {ex.Message}\n\nProgress before failure:\n{_progress}\n\nFull exception:\n{ex}",
                        "Script failed", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            // ------------------------------------------------------------------
            // STEP 1: z_PTV_eval_{dose}_{suffix}
            // ------------------------------------------------------------------
            private void Step1_EvalPtv(
                List<OptKey> groupKeys,
                Dictionary<OptKey, List<Structure>> targetsByGroup,
                Structure ext,
                SegmentVolume bodyMinus3)
            {
                LogSection("1) z_PTV_eval_{dose}_{suffix}");
                foreach (var k in groupKeys)
                {
                    if (!targetsByGroup.TryGetValue(k, out var src) || src == null || src.Count == 0) continue;

                    string doseStr = k.DoseGy.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string sfxStr = string.IsNullOrWhiteSpace(k.Suffix) ? "" : "_" + k.Suffix;
                    string evalId = TruncId($"z_PTV_eval_{doseStr}{sfxStr}");

                    using (var tg = new TempGuard(_ss))
                    {
                        var doseUnionSt = tg.Add(UnionManyToTemp(_ss, src, _fb, "zTmpDoseU",
                                            $"DoseUnion_{doseStr}{sfxStr}", tg));
                        if (doseUnionSt == null) continue;

                        var evalSeg = SafeBoolean(_ss, doseUnionSt.SegmentVolume, bodyMinus3,
                                        BoolOp.And, doseUnionSt, ext, evalId, _fb,
                                        $"Eval_{doseStr}{sfxStr}_AndBodyMinus3", tg);
                        evalSeg = SafeBoolean(_ss, evalSeg, ext.SegmentVolume,
                                    BoolOp.And, null, ext, evalId, _fb,
                                    $"Eval_{doseStr}{sfxStr}_AndExt", tg);

                        if (evalSeg != null)
                        {
                            var st = GetOrCreate(_ss, "PTV", evalId);
                            if (AssignSegmentSafely(st, evalSeg))
                            {
                                st.Color = Colors.Blue;
                                _zEval[k] = st;
                                LogCreated(evalId);
                            }
                            else
                            {
                                _ss.RemoveStructure(st);
                                _progress.AppendLine($"  SKIP: {evalId} (Empty volume)");
                            }
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 2: z_PTV_opt_{dose}_{suffix}
            // ------------------------------------------------------------------
            private void Step2_OptPtv(List<OptKey> groupKeys, Structure ext)
            {
                LogSection("2) z_PTV_opt_{dose}_{suffix}");
                foreach (var k in groupKeys)
                {
                    if (!_zEval.TryGetValue(k, out var evalSt)) continue;

                    string doseStr = k.DoseGy.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string sfxStr = string.IsNullOrWhiteSpace(k.Suffix) ? "" : "_" + k.Suffix;
                    string optId = TruncId($"z_PTV_opt_{doseStr}{sfxStr}");

                    using (var tg = new TempGuard(_ss))
                    {
                        var evalClone = CloneSegViaTempTracked(_ss, evalSt.SegmentVolume, "zTmpEval", tg);
                        var optSeg = SafeMargin(evalClone, +EVAL_TO_OPT_EXPAND_MM);

                        // Subtract all higher-dose opt structures + 1 mm
                        foreach (var hk in groupKeys.Where(x => x.DoseGy > k.DoseGy))
                        {
                            if (!_zOpt.TryGetValue(hk, out var higherOpt)) continue;

                            var higherClone = CloneSegViaTempTracked(_ss, higherOpt.SegmentVolume, "zTmpOptH", tg);
                            var higherExpanded = SafeMargin(higherClone, +LOWER_SUBTRACT_EXTRA_MM);

                            string hDoseStr = hk.DoseGy.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            string hSfxStr = string.IsNullOrWhiteSpace(hk.Suffix) ? "" : "_" + hk.Suffix;

                            optSeg = SafeBoolean(_ss, optSeg, higherExpanded, BoolOp.Sub,
                                        null, higherOpt, optId, _fb,
                                        $"Opt_{doseStr}{sfxStr}_SubHigher_{hDoseStr}{hSfxStr}", tg);
                        }

                        optSeg = SafeBoolean(_ss, optSeg, ext.SegmentVolume, BoolOp.And,
                                    null, ext, optId, _fb, $"Opt_{doseStr}{sfxStr}_CapExt", tg);

                        if (optSeg != null)
                        {
                            var st = GetOrCreate(_ss, "PTV", optId);
                            if (AssignSegmentSafely(st, optSeg))
                            {
                                st.Color = Colors.Red;
                                if (SMOOTH_OPT_TARGET) SmoothStructureByExpandContract(st, SMOOTH_MM);
                                _zOpt[k] = st;
                                LogCreated(optId);
                            }
                            else
                            {
                                _ss.RemoveStructure(st);
                                _progress.AppendLine($"  SKIP: {optId} (Empty volume)");
                            }
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 3a: z_PTV_opt_{dose}_sum
            // ------------------------------------------------------------------
            private void Step3a_DoseOptPtvSum(
                List<OptKey> groupKeys, List<double> doseLevels, Structure ext)
            {
                LogSection("3a) z_PTV_opt_{dose}_sum");
                using (var tg = new TempGuard(_ss))
                {
                    foreach (var d in doseLevels)
                    {
                        var groupsAtDose = groupKeys
                            .Where(k => k.DoseGy == d && _zOpt.ContainsKey(k))
                            .Select(k => _zOpt[k])
                            .ToList();
                        if (groupsAtDose.Count == 0) continue;

                        string doseStr = d.ToString(System.Globalization.CultureInfo.InvariantCulture);
                        string sumId = TruncId($"z_PTV_opt_{doseStr}_sum");

                        var u = tg.Add(UnionManyToTemp(_ss, groupsAtDose, _fb,
                                    "zTmpOptDoseU", $"OptDoseUnion_{doseStr}", tg));
                        if (u == null) continue;

                        var seg = SafeBoolean(_ss, u.SegmentVolume, ext.SegmentVolume, BoolOp.And,
                                    u, ext, sumId, _fb, $"OptDoseSum_{doseStr}_AndExt", tg);
                        if (seg != null)
                        {
                            var st = GetOrCreate(_ss, "PTV", sumId);
                            if (AssignSegmentSafely(st, seg))
                            {
                                st.Color = Colors.Red;
                                if (SMOOTH_OPT_TARGET) SmoothStructureByExpandContract(st, SMOOTH_MM);
                                _zOptDoseSum[d] = st;
                                LogCreated(sumId);
                            }
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 3b: z_PTV_opt_sum (global)
            // ------------------------------------------------------------------
            private void Step3b_GlobalOptPtvSum(List<Structure> selectedTargets, Structure ext)
            {
                LogSection("3b) z_PTV_opt_sum");
                using (var tg = new TempGuard(_ss))
                {
                    var tvUnionSt = tg.Add(UnionManyToTemp(_ss, selectedTargets, _fb,
                                        "zTmpAllTvU", "AllTvUnion", tg));
                    if (tvUnionSt == null) return;

                    var optSumSeg = SafeMargin(tvUnionSt.SegmentVolume, +EVAL_TO_OPT_EXPAND_MM);
                    optSumSeg = SafeBoolean(_ss, optSumSeg, ext.SegmentVolume, BoolOp.And,
                                    null, ext, ID_OPT_TV_SUM, _fb, "OptSumAndExt", tg);

                    if (optSumSeg != null)
                    {
                        var zOptSum = GetOrCreate(_ss, "PTV", ID_OPT_TV_SUM);
                        if (AssignSegmentSafely(zOptSum, optSumSeg))
                        {
                            zOptSum.Color = Colors.Red;
                            if (SMOOTH_OPT_TARGET) SmoothStructureByExpandContract(zOptSum, SMOOTH_MM);
                            LogCreated(ID_OPT_TV_SUM);
                        }
                        else
                        {
                            _ss.RemoveStructure(zOptSum);
                            _progress.AppendLine($"  SKIP: {ID_OPT_TV_SUM} (Empty volume)");
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 4: Avoidance structures
            // ------------------------------------------------------------------
            private void Step4_Avoidance(
                List<OptKey> groupKeys,
                List<double> doseLevels,
                List<TargetDosePair> targetDosePairs,
                Structure ext)
            {
                LogSection("4a) zAvoidance_{dose} (Automatic Dose-Level)");
                foreach (var d in doseLevels)
                {
                    if (!_zOptDoseSum.TryGetValue(d, out var optSumSt)) continue;

                    string doseStr = d.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string avoidId = TruncId($"zAvoidance_{doseStr}");

                    using (var tg = new TempGuard(_ss))
                    {
                        var expanded = SafeMargin(optSumSt.SegmentVolume, +AVOIDANCE_MARGIN_MM);
                        var avoidanceSeg = SafeBoolean(_ss, ext.SegmentVolume, expanded, BoolOp.Sub,
                                            ext, optSumSt, avoidId, _fb,
                                            $"Avoid_{doseStr}_ExtMinusExpOpt", tg);
                        avoidanceSeg = SafeBoolean(_ss, avoidanceSeg, ext.SegmentVolume, BoolOp.And,
                                        null, ext, avoidId, _fb,
                                        $"Avoid_{doseStr}_CapExt", tg);

                        if (avoidanceSeg != null)
                        {
                            var st = GetOrCreate(_ss, "CONTROL", avoidId);
                            if (AssignSegmentSafely(st, avoidanceSeg)) LogCreated(avoidId);
                            else { _ss.RemoveStructure(st); _progress.AppendLine($"  SKIP: {avoidId} (Empty volume)"); }
                        }
                    }
                }

                LogSection("4b) zAvoidance_{dose}_{suffix} (Target-Specific Overrides)");
                foreach (var k in groupKeys)
                {
                    bool anyAvoid = targetDosePairs.Any(t =>
                        t.DoseGy == k.DoseGy &&
                        string.Equals(t.Suffix, k.Suffix, StringComparison.OrdinalIgnoreCase) &&
                        t.CreateAvoidance);
                    if (!anyAvoid) continue;
                    if (!_zOpt.TryGetValue(k, out var optSt)) continue;

                    string doseStr = k.DoseGy.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string sfxStr = string.IsNullOrWhiteSpace(k.Suffix) ? "" : "_" + k.Suffix;
                    string avoidId = TruncId($"zAvoidance_{doseStr}{sfxStr}");

                    using (var tg = new TempGuard(_ss))
                    {
                        var expanded = SafeMargin(optSt.SegmentVolume, +AVOIDANCE_MARGIN_MM);
                        var avoidanceSeg = SafeBoolean(_ss, ext.SegmentVolume, expanded, BoolOp.Sub,
                                            ext, optSt, avoidId, _fb,
                                            $"Avoid_{doseStr}{sfxStr}_ExtMinusExpOpt", tg);
                        avoidanceSeg = SafeBoolean(_ss, avoidanceSeg, ext.SegmentVolume, BoolOp.And,
                                        null, ext, avoidId, _fb,
                                        $"Avoid_{doseStr}{sfxStr}_CapExt", tg);

                        if (avoidanceSeg != null)
                        {
                            var st = GetOrCreate(_ss, "CONTROL", avoidId);
                            if (AssignSegmentSafely(st, avoidanceSeg)) LogCreated(avoidId);
                            else { _ss.RemoveStructure(st); _progress.AppendLine($"  SKIP: {avoidId} (Empty volume)"); }
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 5: Virtual Bolus pipeline
            //
            // v3.0.0.30: physicalBolus parameter added.
            // When non-null, the skin baseline for all Step 5 operations is
            //   Body_with_Bolus = original_body Or physicalBolus
            // instead of original_body. This places the virtual bolus on top of
            // the physical bolus surface. Steps 6-9 are unaffected.
            //
            // Pipeline (physical bolus present):
            //   0. Body_with_Bolus (temp) = body Or physicalBolus
            //   1. z_Virtual_PTV  = z_PTV_opt expanded ant+lat, lung subtracted
            //                       (same as before, but relative to Body_with_Bolus)
            //   2. zVB_RawBolus (temp) = z_Virtual_PTV Sub Body_with_Bolus
            //   3. Body_new = Body_with_Bolus Or z_Virtual_PTV  [CONTROL structure]
            //   4. z_Virtual_Bolus = rawBolus And SafeMargin(Body_new, -VB_SKIN_CROP_MM)
            //   5. z_Virtual_PTV_Opt = z_Virtual_Bolus expanded post+inf VB_OPT_INWARD_MM,
            //                          capped to SafeMargin(Body_new, -VB_OPT_SKIN_CROP_MM)
            //
            // Pipeline (no physical bolus – original behaviour):
            //   skinRef = original_body (ext) throughout.
            // ------------------------------------------------------------------
            private Structure Step5_VirtualBolus(
                List<TargetDosePair> bolusRequests,
                Structure ext,
                bool isLeft,
                TempGuard globalTg,
                Structure physicalBolus,        // NEW PARAMETER (v3.0.0.30)
                double physicalThicknessMm)     // NEW PARAMETER (v3.0.0.32)
            {
                if (bolusRequests == null || bolusRequests.Count == 0) return null;
                LogSection("5) Virtual Bolus Pipeline");

                using (var tg = new TempGuard(_ss))
                {
                    // ============================================================
                    // STEP 0 (v3.0.0.30): Build Body_with_Bolus when physical bolus present.
                    // This becomes the skin reference for all of Step 5.
                    // When no physical bolus, skinRefSt = ext (original behaviour).
                    // ============================================================
                    Structure skinRefSt = ext;  // default: original body

                    if (physicalBolus != null)
                    {
                        LogSection("5.0) Body_with_Bolus = body Or physicalBolus");

                        var origBodyForBwB = tg.Add(_ss.AddStructure("CONTROL",
                            MakeUniqueId(_ss, "zVB_BodyBolusTmp")));
                        if (ext.IsHighResolution && !origBodyForBwB.IsHighResolution)
                            origBodyForBwB.ConvertToHighResolution();
                        AssignSegmentSafely(origBodyForBwB, ext.SegmentVolume);

                        var physBolusTmp = tg.Add(_ss.AddStructure("CONTROL",
                            MakeUniqueId(_ss, "zVB_PhysBolusTmp")));
                        if (physicalBolus.IsHighResolution && !physBolusTmp.IsHighResolution)
                            physBolusTmp.ConvertToHighResolution();
                        AssignSegmentSafely(physBolusTmp, physicalBolus.SegmentVolume);

                        var bwbSeg = SafeBoolean(_ss,
                            origBodyForBwB.SegmentVolume, physBolusTmp.SegmentVolume,
                            BoolOp.Or, origBodyForBwB, physBolusTmp,
                            "Body_with_Bolus", _fb, "VB_BodyWithBolus_OrPhysBolus", tg);

                        if (bwbSeg != null)
                        {
                            var bwbSt = tg.Add(_ss.AddStructure("CONTROL",
                                MakeUniqueId(_ss, "zVB_BodyWBolus")));
                            if (origBodyForBwB.IsHighResolution && !bwbSt.IsHighResolution)
                                bwbSt.ConvertToHighResolution();
                            if (AssignSegmentSafely(bwbSt, bwbSeg))
                            {
                                skinRefSt = bwbSt;
                                _progress.AppendLine($"  Body_with_Bolus created from body Or {physicalBolus.Id}");
                            }
                            else
                            {
                                _progress.AppendLine("  WARN (5.0): Body_with_Bolus assignment failed – falling back to original body.");
                            }
                        }
                        else
                        {
                            _progress.AppendLine("  WARN (5.0): Body_with_Bolus boolean failed – falling back to original body.");
                        }
                    }

                    // ============================================================
                    // STEP 1: z_Virtual_PTV
                    //   Union of all z_PTV_opt targets, each expanded ant + lateral
                    //   by totalMm, then subtract ipsilateral lung.
                    //   totalMm = bolusMm + physicalThicknessMm when Physical Bolus is
                    //   ticked (v3.0.0.32; no extra buffer), otherwise
                    //   bolusMm + VB_EXTRA_EXPAND_MM (original behaviour).
                    //   Uses skinRefSt (Body_with_Bolus when physical bolus present).
                    // ============================================================
                    var ipsiLung = FindIpsilateralLung(_ss, isLeft);
                    Structure ipsiLungSt = null;
                    if (ipsiLung != null)
                    {
                        ipsiLungSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_Lung")));
                        if (ipsiLung.IsHighResolution && !ipsiLungSt.IsHighResolution)
                            ipsiLungSt.ConvertToHighResolution();
                        AssignSegmentSafely(ipsiLungSt, ipsiLung.SegmentVolume);
                        _progress.AppendLine($"  Lung subtraction: {ipsiLung.Id}");
                    }
                    else
                    {
                        _progress.AppendLine("  WARNING: No ipsilateral lung found – lung subtraction skipped.");
                    }

                    Structure vPtvAccSt = null;

                    foreach (var req in bolusRequests)
                    {
                        var k = new OptKey(req.DoseGy, req.Suffix);
                        if (!_zOpt.ContainsKey(k)) continue;

                        double totalMm = physicalBolus != null
                            ? req.BolusMm.GetValueOrDefault() + physicalThicknessMm
                            : req.BolusMm.GetValueOrDefault() + VB_EXTRA_EXPAND_MM;

                        var baseSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_Base")));
                        AssignSegmentSafely(baseSt, _zOpt[k].SegmentVolume);

                        // Subtract ipsilateral lung from base BEFORE expansion
                        if (ipsiLungSt != null && !ipsiLungSt.IsEmpty)
                        {
                            var subSeg = SafeBoolean(_ss, baseSt.SegmentVolume, ipsiLungSt.SegmentVolume,
                                            BoolOp.Sub, baseSt, ipsiLungSt, null, _fb,
                                            $"VB_Base_SubLung_{req.TargetId}", tg);
                            if (subSeg != null) AssignSegmentSafely(baseSt, subSeg);
                        }

                        var expSeg = SafePerpendicularMargin(_ss, baseSt, ext, totalMm,
                                        isLeft, _fb, tg, $"zVB_AsymExp_{req.TargetId}");
                        if (expSeg == null) continue;

                        // Subtract lung from expanded result too (catches margin bleed-in)
                        if (ipsiLungSt != null && !ipsiLungSt.IsEmpty)
                        {
                            var expSt2 = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_ExpTmp")));
                            AssignSegmentSafely(expSt2, expSeg);
                            var subSeg2 = SafeBoolean(_ss, expSt2.SegmentVolume, ipsiLungSt.SegmentVolume,
                                            BoolOp.Sub, expSt2, ipsiLungSt, null, _fb,
                                            $"VB_Exp_SubLung_{req.TargetId}", tg);
                            if (subSeg2 != null) expSeg = subSeg2;
                        }

                        var expSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_Exp")));
                        AssignSegmentSafely(expSt, expSeg);

                        if (vPtvAccSt == null)
                        {
                            vPtvAccSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_PtvAcc")));
                            AssignSegmentSafely(vPtvAccSt, expSt.SegmentVolume);
                        }
                        else
                        {
                            var unionSeg = SafeBoolean(_ss, vPtvAccSt.SegmentVolume, expSt.SegmentVolume,
                                            BoolOp.Or, vPtvAccSt, expSt, null, _fb,
                                            $"VB_PtvUnion_{req.TargetId}", tg);
                            if (unionSeg != null) AssignSegmentSafely(vPtvAccSt, unionSeg);
                        }
                    }

                    if (vPtvAccSt == null || vPtvAccSt.IsEmpty)
                    {
                        _progress.AppendLine("  SKIP: Virtual Bolus (no valid opt targets with bolus input)");
                        return null;
                    }

                    // Write permanent z_Virtual_PTV structure
                    var zVirtualPtv = GetOrCreate(_ss, "PTV", "z_Virtual_PTV");
                    if (!AssignSegmentSafely(zVirtualPtv, vPtvAccSt.SegmentVolume))
                    {
                        _ss.RemoveStructure(zVirtualPtv);
                        _progress.AppendLine("  SKIP: z_Virtual_PTV (empty after expansion)");
                        return null;
                    }
                    zVirtualPtv.Color = Color.FromRgb(160, 32, 240);
                    LogCreated("z_Virtual_PTV");

                    // ============================================================
                    // STEP 1b (v3.0.0.31): Physical bolus optimisation structures.
                    //   Bolus_physical = copy of the physical bolus (type CONTROL).
                    //   Bolus_phys_Opt  = overlap of Bolus_physical and z_Virtual_PTV
                    //                     (type CONTROL, same as Bolus_physical).
                    //   Only created when Physical Bolus is ticked (physicalBolus != null).
                    //
                    //   NOTE: Eclipse's scripting API refuses AddStructure(dicomType:
                    //   "BOLUS", ...) - "Can not add a new structure: invalid DICOM
                    //   type string" - BOLUS structures can only be created via
                    //   Insert > New Bolus... in the Eclipse UI, never by script (this
                    //   is the same reason physicalBolus itself must already exist).
                    //   Bolus_physical is therefore type CONTROL, matching Body_new
                    //   and the other script-created helper structures.
                    //
                    //   bolusPhysical is hoisted to method scope (assigned here, used
                    //   again below in the physical-bolus z_Virtual_Bolus formula).
                    // ============================================================
                    Structure bolusPhysical = null;
                    if (physicalBolus != null)
                    {
                        bolusPhysical = GetOrCreate(_ss, "CONTROL", "Bolus_physical");
                        if (physicalBolus.IsHighResolution && !bolusPhysical.IsHighResolution)
                            bolusPhysical.ConvertToHighResolution();

                        bool bolusPhysicalOk;
                        string bolusPhysicalFailReason = null;
                        try
                        {
                            bolusPhysical.SegmentVolume = physicalBolus.SegmentVolume;
                            bolusPhysicalOk = !bolusPhysical.IsEmpty;
                            if (!bolusPhysicalOk)
                                bolusPhysicalFailReason =
                                    $"assignment succeeded but result is empty (source {physicalBolus.Id}.SegmentVolume may itself be empty/unset)";
                        }
                        catch (Exception exCopy)
                        {
                            bolusPhysicalOk = false;
                            bolusPhysicalFailReason = exCopy.Message;
                        }

                        if (bolusPhysicalOk)
                        {
                            bolusPhysical.Color = Color.FromRgb(255, 165, 0);
                            LogCreated("Bolus_physical");

                            var bolusPhysOptSeg = SafeBoolean(_ss,
                                bolusPhysical.SegmentVolume, zVirtualPtv.SegmentVolume,
                                BoolOp.And, bolusPhysical, zVirtualPtv, "Bolus_phys_Opt", _fb,
                                "VB_BolusPhysOpt_AndVirtualPtv", tg);

                            if (bolusPhysOptSeg != null)
                            {
                                var bolusPhysOpt = GetOrCreate(_ss, "CONTROL", "Bolus_phys_Opt");
                                if (AssignSegmentSafely(bolusPhysOpt, bolusPhysOptSeg))
                                {
                                    bolusPhysOpt.Color = Color.FromRgb(255, 200, 0);
                                    LogCreated("Bolus_phys_Opt");
                                }
                                else
                                {
                                    _ss.RemoveStructure(bolusPhysOpt);
                                    _progress.AppendLine("  SKIP: Bolus_phys_Opt (no overlap with z_Virtual_PTV)");
                                }
                            }
                            else
                            {
                                _progress.AppendLine("  SKIP: Bolus_phys_Opt (boolean failed)");
                            }
                        }
                        else
                        {
                            _ss.RemoveStructure(bolusPhysical);
                            _progress.AppendLine($"  WARN (1b): Bolus_physical copy failed - {bolusPhysicalFailReason}");
                        }
                    }

                    // ============================================================
                    // STEP 2/3 (common to both modes): skin snapshot + Body_new.
                    //   Body_new = skinRefSt Or z_Virtual_PTV  [CONTROL structure]
                    //   When physical bolus is present: skinRefSt = Body_with_Bolus,
                    //   so Body_new = (body Or physBolus) Or z_Virtual_PTV.
                    //   When no physical bolus: Body_new = body Or z_Virtual_PTV (original).
                    // ============================================================
                    var skinSnapSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_SkinSnap")));
                    if (skinRefSt.IsHighResolution && !skinSnapSt.IsHighResolution)
                        skinSnapSt.ConvertToHighResolution();
                    AssignSegmentSafely(skinSnapSt, skinRefSt.SegmentVolume);

                    var bodyNewSeg = SafeBoolean(_ss, skinSnapSt.SegmentVolume, zVirtualPtv.SegmentVolume,
                                        BoolOp.Or, skinSnapSt, zVirtualPtv, "Body_new", _fb,
                                        "VB_BodyNew_OrVPtv", tg);

                    Structure bodyNew = null;
                    if (bodyNewSeg != null)
                    {
                        bodyNew = GetOrCreate(_ss, "CONTROL", "Body_new");
                        if (AssignSegmentSafely(bodyNew, bodyNewSeg))
                        {
                            bodyNew.Color = Color.FromRgb(255, 255, 128);
                            LogCreated("Body_new");
                        }
                        else
                        {
                            _ss.RemoveStructure(bodyNew);
                            bodyNew = null;
                            _progress.AppendLine("  WARN (3): Body_new could not be assigned.");
                        }
                    }
                    else
                    {
                        _progress.AppendLine("  WARN (3): Body_new boolean failed.");
                    }

                    if (physicalBolus != null)
                    {
                        // ========================================================
                        // STEP 4-5 (v3.0.0.32, physical bolus mode):
                        //   z_Virtual_Bolus   = z_Virtual_PTV Sub Bolus_physical Sub Body
                        //   z_Virtual_PTV_Opt = (union of z_PTV_opt targets feeding
                        //                        z_Virtual_PTV) expanded ant+lat by
                        //                        VB_PHYS_OPT_EXPAND_MM, capped to Body_new.
                        // ========================================================
                        var afterPhysBolusSeg = (bolusPhysical != null && !bolusPhysical.IsEmpty)
                            ? SafeBoolean(_ss, zVirtualPtv.SegmentVolume, bolusPhysical.SegmentVolume,
                                  BoolOp.Sub, zVirtualPtv, bolusPhysical, null, _fb,
                                  "VB_Bolus_SubPhysBolus", tg)
                            : zVirtualPtv.SegmentVolume;

                        var virtualBolusSeg = SafeBoolean(_ss, afterPhysBolusSeg, ext.SegmentVolume,
                                                BoolOp.Sub, null, ext, null, _fb,
                                                "VB_Bolus_SubBody", tg);

                        if (virtualBolusSeg == null)
                        {
                            _progress.AppendLine(
                                "  SKIP: z_Virtual_Bolus (empty after subtracting Bolus_physical + Body)");
                        }
                        else
                        {
                            var zVirtualBolusPhys = GetOrCreate(_ss, "PTV", "z_Virtual_Bolus");
                            if (AssignSegmentSafely(zVirtualBolusPhys, virtualBolusSeg))
                            {
                                zVirtualBolusPhys.Color = Color.FromRgb(160, 32, 240);
                                LogCreated("z_Virtual_Bolus");
                            }
                            else
                            {
                                _ss.RemoveStructure(zVirtualBolusPhys);
                                _progress.AppendLine("  SKIP: z_Virtual_Bolus (empty)");
                            }
                        }

                        Structure ptvOptAccSt = null;
                        foreach (var req in bolusRequests)
                        {
                            var k = new OptKey(req.DoseGy, req.Suffix);
                            if (!_zOpt.ContainsKey(k)) continue;

                            var expSeg2 = SafePerpendicularMargin(_ss, _zOpt[k], ext,
                                            VB_PHYS_OPT_EXPAND_MM, isLeft, _fb, tg,
                                            $"zVB_PhysOptExp_{req.TargetId}");
                            if (expSeg2 == null) continue;

                            if (ptvOptAccSt == null)
                            {
                                ptvOptAccSt = tg.Add(_ss.AddStructure("CONTROL",
                                    MakeUniqueId(_ss, "zVB_PhysOptAcc")));
                                AssignSegmentSafely(ptvOptAccSt, expSeg2);
                            }
                            else
                            {
                                var unionSeg2 = SafeBoolean(_ss, ptvOptAccSt.SegmentVolume, expSeg2,
                                                    BoolOp.Or, ptvOptAccSt, null, null, _fb,
                                                    $"VB_PhysOptUnion_{req.TargetId}", tg);
                                if (unionSeg2 != null) AssignSegmentSafely(ptvOptAccSt, unionSeg2);
                            }
                        }

                        if (ptvOptAccSt == null || ptvOptAccSt.IsEmpty)
                        {
                            _progress.AppendLine("  SKIP: z_Virtual_PTV_Opt (no valid PTV_opt targets to expand)");
                            return bodyNew;
                        }

                        var vPtvOptSegPhys = bodyNew != null
                            ? SafeBoolean(_ss, ptvOptAccSt.SegmentVolume, bodyNew.SegmentVolume,
                                  BoolOp.And, ptvOptAccSt, bodyNew, "z_Virtual_PTV_Opt", _fb,
                                  "VBPhysOpt_CapBodyNew", tg)
                            : ptvOptAccSt.SegmentVolume;

                        if (vPtvOptSegPhys == null)
                        {
                            _progress.AppendLine("  SKIP: z_Virtual_PTV_Opt (empty after capping to Body_new)");
                            return bodyNew;
                        }

                        var zVPtvOptPhys = GetOrCreate(_ss, "PTV", "z_Virtual_PTV_Opt");
                        if (AssignSegmentSafely(zVPtvOptPhys, vPtvOptSegPhys))
                        {
                            zVPtvOptPhys.Color = Color.FromRgb(200, 80, 255);
                            LogCreated("z_Virtual_PTV_Opt");
                        }
                        else
                        {
                            _ss.RemoveStructure(zVPtvOptPhys);
                            _progress.AppendLine("  SKIP: z_Virtual_PTV_Opt (empty)");
                        }

                        return bodyNew;
                    }

                    // ============================================================
                    // STEP 2 (no physical bolus): z_Virtual_Bolus_raw (temp)
                    //   = z_Virtual_PTV Sub skinRefSt (skinRefSt = original body)
                    // ============================================================
                    var rawBolusSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_RawBolus")));
                    var rawBolusSeg = SafeBoolean(_ss, zVirtualPtv.SegmentVolume, skinSnapSt.SegmentVolume,
                                        BoolOp.Sub, zVirtualPtv, skinSnapSt, null, _fb,
                                        "VB_RawBolus_SubSkinRef", tg);
                    if (rawBolusSeg != null) AssignSegmentSafely(rawBolusSt, rawBolusSeg);

                    if (rawBolusSt.IsEmpty)
                    {
                        _progress.AppendLine("  NOTE: Virtual_PTV is fully inside the skin reference – bolus will be empty.");
                        return bodyNew;
                    }

                    // ============================================================
                    // STEP 4 (no physical bolus): z_Virtual_Bolus
                    //   = rawBolus And SafeMargin(Body_new, -VB_SKIN_CROP_MM)
                    // ============================================================
                    var finalSkinRef = bodyNew ?? skinRefSt;

                    var bodyMinus2St = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_BdyM2")));
                    if (finalSkinRef.IsHighResolution && !bodyMinus2St.IsHighResolution)
                        bodyMinus2St.ConvertToHighResolution();
                    AssignSegmentSafely(bodyMinus2St, SafeMargin(finalSkinRef.SegmentVolume, -VB_SKIN_CROP_MM));

                    var bolusSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_Bolus")));
                    var bolusSeg = SafeBoolean(_ss, rawBolusSt.SegmentVolume, bodyMinus2St.SegmentVolume,
                                    BoolOp.And, rawBolusSt, bodyMinus2St, null, _fb,
                                    "VB_Bolus_CapBodyMinus2", tg);
                    if (bolusSeg != null) AssignSegmentSafely(bolusSt, bolusSeg);

                    if (bolusSt.IsEmpty)
                    {
                        _progress.AppendLine("  SKIP: z_Virtual_Bolus (empty after 2 mm skin trim)");
                        return bodyNew;
                    }

                    var zVirtualBolus = GetOrCreate(_ss, "PTV", "z_Virtual_Bolus");
                    if (!AssignSegmentSafely(zVirtualBolus, bolusSt.SegmentVolume))
                    {
                        _ss.RemoveStructure(zVirtualBolus);
                        _progress.AppendLine("  SKIP: z_Virtual_Bolus (empty)");
                        return bodyNew;
                    }
                    zVirtualBolus.Color = Color.FromRgb(160, 32, 240);
                    LogCreated("z_Virtual_Bolus");

                    // ============================================================
                    // STEP 5 (no physical bolus): z_Virtual_PTV_Opt
                    //   = z_Virtual_Bolus expanded post+inf VB_OPT_INWARD_MM,
                    //     capped to SafeMargin(Body_new, -VB_OPT_SKIN_CROP_MM)
                    // ============================================================
                    var vPtvOptMargins = new AxisAlignedMargins(
                        StructureMarginGeometry.Outer,
                        0,                  // x1 Right    – no expansion
                        0,                  // y1 Anterior – no expansion
                        VB_OPT_INWARD_MM,   // z1 Inferior – push toward feet
                        0,                  // x2 Left     – no expansion
                        VB_OPT_INWARD_MM,   // y2 Posterior – push toward spine
                        0);                 // z2 Superior – no expansion

                    SegmentVolume vPtvOptExpandedSeg;
                    try
                    {
                        vPtvOptExpandedSeg = zVirtualBolus.SegmentVolume.AsymmetricMargin(vPtvOptMargins);
                    }
                    catch
                    {
                        _fb.MarkCreated("z_Virtual_PTV_Opt_IsoFallback");
                        vPtvOptExpandedSeg = SafeMargin(zVirtualBolus.SegmentVolume, VB_OPT_INWARD_MM);
                    }

                    var vPtvOptExpSt = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_OptExp")));
                    AssignSegmentSafely(vPtvOptExpSt, vPtvOptExpandedSeg);

                    var bodyMinus4St = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zVB_BdyM4")));
                    if (finalSkinRef.IsHighResolution && !bodyMinus4St.IsHighResolution)
                        bodyMinus4St.ConvertToHighResolution();
                    AssignSegmentSafely(bodyMinus4St, SafeMargin(finalSkinRef.SegmentVolume, -VB_OPT_SKIN_CROP_MM));

                    var vPtvOptSeg = SafeBoolean(_ss, vPtvOptExpSt.SegmentVolume, bodyMinus4St.SegmentVolume,
                                        BoolOp.And, vPtvOptExpSt, bodyMinus4St, "z_Virtual_PTV_Opt",
                                        _fb, "VBOpt_CapBodyMinus4", tg);

                    if (vPtvOptSeg == null)
                    {
                        _progress.AppendLine("  SKIP: z_Virtual_PTV_Opt (empty after 4mm skin-crop cap)");
                        return bodyNew;
                    }

                    var zVPtvOpt = GetOrCreate(_ss, "PTV", "z_Virtual_PTV_Opt");
                    if (AssignSegmentSafely(zVPtvOpt, vPtvOptSeg))
                    {
                        zVPtvOpt.Color = Color.FromRgb(200, 80, 255);
                        LogCreated("z_Virtual_PTV_Opt");
                    }
                    else
                    {
                        _ss.RemoveStructure(zVPtvOpt);
                        _progress.AppendLine("  SKIP: z_Virtual_PTV_Opt (empty)");
                    }

                    return bodyNew;
                }
            }


            // ------------------------------------------------------------------
            // STEP 6: z_[OAR]_Ovl_[dose]
            // ------------------------------------------------------------------
            private void Step6_Overlaps(
                List<OrganRow> organRows, List<double> doseLevels, Structure ext)
            {
                LogSection("6) z_[OAR]_Ovl_[dose]");
                var ovlRequests = organRows.Where(r => r.CreateOvl).ToList();
                _progress.AppendLine($"  ({ovlRequests.Count} Ovl requests x {doseLevels.Count} dose levels)");

                foreach (var d in doseLevels)
                {
                    if (!_zOptDoseSum.TryGetValue(d, out var optSumSt)) continue;

                    string doseStr = d.ToString(System.Globalization.CultureInfo.InvariantCulture);

                    foreach (var req in ovlRequests)
                    {
                        string ovlDisplayName = $"z_{req.OarId}_Ovl_{doseStr}";
                        try
                        {
                            var oar = _ss.Structures.FirstOrDefault(s =>
                                !s.IsEmpty && string.Equals(s.Id, req.OarId, StringComparison.OrdinalIgnoreCase));
                            if (oar == null || oar.IsEmpty)
                            {
                                _progress.AppendLine($"  SKIP: {ovlDisplayName} (OAR not found or empty)");
                                continue;
                            }

                            using (var tg = new TempGuard(_ss))
                            {
                                var optSeg = CloneSegViaTempTracked(_ss, optSumSt.SegmentVolume, "zTmpOptOvlp", tg);
                                var oarSeg = CloneSegViaTempTracked(_ss, oar.SegmentVolume, "zTmpOarOvlp", tg);
                                var ovlId = BuildId("z_", req.OarId, $"_Ovl_{doseStr}");
                                var opToUse = OVERLAP_IS_INTERSECTION ? BoolOp.And : BoolOp.Or;
                                var actionName = OVERLAP_IS_INTERSECTION ? "And" : "Or";

                                var ov = SafeBoolean(_ss, optSeg, oarSeg, opToUse,
                                            optSumSt, oar, ovlId, _fb,
                                            $"Ovlp_{doseStr}_{req.OarId}_{actionName}", tg);

                                // Crop lower-dose overlap from higher-dose opt targets + 1 mm
                                foreach (var hd in doseLevels.Where(x => x > d))
                                {
                                    if (!_zOptDoseSum.TryGetValue(hd, out var higherOptSum)) continue;

                                    var higherClone = CloneSegViaTempTracked(_ss, higherOptSum.SegmentVolume, "zTmpOptH", tg);
                                    var higherExpanded = SafeMargin(higherClone, +LOWER_SUBTRACT_EXTRA_MM);
                                    string hDoseStr = hd.ToString(System.Globalization.CultureInfo.InvariantCulture);

                                    ov = SafeBoolean(_ss, ov, higherExpanded, BoolOp.Sub,
                                            null, higherOptSum, ovlId, _fb,
                                            $"Ovlp_{doseStr}_SubHigher_{hDoseStr}", tg);
                                }

                                ov = SafeBoolean(_ss, ov, ext.SegmentVolume, BoolOp.And,
                                        null, ext, ovlId, _fb,
                                        $"Ovlp_{doseStr}_{req.OarId}_CapExt", tg);

                                if (ov != null)
                                {
                                    var st = GetOrCreate(_ss, "PTV", ovlId);
                                    if (AssignSegmentSafely(st, ov))
                                    {
                                        st.Color = Color.FromRgb(191, 255, 0);
                                        LogCreated(ovlId);
                                    }
                                    else
                                    {
                                        _ss.RemoveStructure(st);
                                        _progress.AppendLine($"  SKIP: {ovlId} (No overlap intersection)");
                                    }
                                }
                            }
                        }
                        catch (Exception exOvl)
                        {
                            _progress.AppendLine($"  FAIL: {ovlDisplayName} -> {exOvl.Message}");
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 7: z_[OAR]_Opt
            // ------------------------------------------------------------------
            private void Step7_OptOars(
                List<OrganRow> organRows, List<OptKey> groupKeys, Structure ext)
            {
                LogSection("7) z_[OAR]_Opt");
                var optRequests = organRows.Where(r => r.CreateOpt).ToList();
                _progress.AppendLine($"  ({optRequests.Count} Opt requests)");

                if (optRequests.Count == 0) return;

                using (var tg = new TempGuard(_ss))
                {
                    SegmentVolume evalPlus2UnionSeg = null;
                    Structure evalPlus2Acc = null;

                    foreach (var k in groupKeys)
                    {
                        if (!_zEval.TryGetValue(k, out var evalSt)) continue;

                        var evalClone2 = CloneSegViaTempTracked(_ss, evalSt.SegmentVolume, "zTmpEval2", tg);
                        var plus2 = SafeMargin(evalClone2, +EVAL_TO_OPT_EXPAND_MM);
                        plus2 = SmoothSegByExpandContract(plus2, SMOOTH_MM);

                        var plus2St = tg.Add(_ss.AddStructure("CONTROL", MakeUniqueId(_ss, "zRC_Plus2")));
                        try { plus2St.SegmentVolume = plus2; }
                        catch { plus2St.ConvertToHighResolution(); plus2St.SegmentVolume = plus2; }

                        if (evalPlus2Acc == null)
                        {
                            evalPlus2Acc = plus2St;
                        }
                        else
                        {
                            var unionSeg = SafeBoolean(_ss, evalPlus2Acc.SegmentVolume,
                                                plus2St.SegmentVolume, BoolOp.Or,
                                                evalPlus2Acc, plus2St, null, _fb, "EvalPlus2Union", tg);
                            if (unionSeg != null)
                                AssignSegmentSafely(evalPlus2Acc, unionSeg);
                        }
                    }

                    if (evalPlus2Acc != null)
                    {
                        evalPlus2UnionSeg = SafeBoolean(_ss, evalPlus2Acc.SegmentVolume,
                                                ext.SegmentVolume, BoolOp.And,
                                                evalPlus2Acc, ext, null, _fb,
                                                "EvalPlus2_CapExt", tg);
                    }

                    foreach (var req in optRequests)
                    {
                        try
                        {
                            var oar = _ss.Structures.FirstOrDefault(s =>
                                !s.IsEmpty && string.Equals(s.Id, req.OarId, StringComparison.OrdinalIgnoreCase));
                            if (oar == null || oar.IsEmpty)
                            {
                                _progress.AppendLine($"  SKIP: z_{req.OarId}_Opt (OAR not found or empty)");
                                continue;
                            }

                            using (var innerTg = new TempGuard(_ss))
                            {
                                var seg = CloneSegViaTempTracked(_ss, oar.SegmentVolume, "zTmpOar", innerTg);
                                var optId = BuildId("z_", req.OarId, "_Opt");

                                if (evalPlus2UnionSeg != null)
                                    seg = SafeBoolean(_ss, seg, evalPlus2UnionSeg, BoolOp.Sub,
                                            oar, null, optId, _fb,
                                            $"OptOAR_{req.OarId}_SubEvalPlus2", innerTg);

                                seg = SafeBoolean(_ss, seg, ext.SegmentVolume, BoolOp.And,
                                        null, ext, optId, _fb,
                                        $"OptOAR_{req.OarId}_CapExt", innerTg);

                                if (seg != null)
                                {
                                    var st = GetOrCreate(_ss, "ORGAN", optId);
                                    if (AssignSegmentSafely(st, seg))
                                    {
                                        st.Color = Color.FromRgb(0, 191, 255);
                                        LogCreated(optId);
                                    }
                                    else
                                    {
                                        _ss.RemoveStructure(st);
                                        _progress.AppendLine($"  SKIP: {optId} (Empty volume post-crop)");
                                    }
                                }
                            }
                        }
                        catch (Exception exOpt)
                        {
                            _progress.AppendLine($"  FAIL: z_{req.OarId}_Opt -> {exOpt.Message}");
                        }
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 8: PRV_[OAR]
            // ------------------------------------------------------------------
            private void Step8_Prvs(List<OrganRow> organRows, Structure ext)
            {
                LogSection("8) PRV_[OAR]");
                var prvRequests = organRows
                    .Where(r => r.CreatePrv && r.ParsedPrvMarginMm.GetValueOrDefault() > 0)
                    .ToList();
                _progress.AppendLine($"  ({prvRequests.Count} PRV requests)");

                foreach (var req in prvRequests)
                {
                    try
                    {
                        var oar = _ss.Structures.FirstOrDefault(s =>
                            !s.IsEmpty && string.Equals(s.Id, req.OarId, StringComparison.OrdinalIgnoreCase));
                        if (oar == null || oar.IsEmpty)
                        {
                            _progress.AppendLine($"  SKIP: PRV_{req.OarId} (OAR not found or empty)");
                            continue;
                        }

                        using (var tg = new TempGuard(_ss))
                        {
                            var oarTemp = tg.Add(_ss.AddStructure("CONTROL",
                                MakeUniqueId(_ss, TruncId("zTmpPrvOar"))));
                            if (oar.IsHighResolution && !oarTemp.IsHighResolution)
                                oarTemp.ConvertToHighResolution();
                            AssignSegmentSafely(oarTemp, oar.SegmentVolume);

                            double marginMm = req.ParsedPrvMarginMm.GetValueOrDefault();
                            var prvSeg = SafeMargin(oarTemp.SegmentVolume, marginMm);

                            var extTemp = tg.Add(_ss.AddStructure("CONTROL",
                                MakeUniqueId(_ss, TruncId("zTmpPrvExt"))));
                            if (oarTemp.IsHighResolution && !extTemp.IsHighResolution)
                                extTemp.ConvertToHighResolution();

                            try
                            {
                                AssignSegmentSafely(extTemp, ext.SegmentVolume);
                            }
                            catch
                            {
                                if (!oarTemp.IsHighResolution) oarTemp.ConvertToHighResolution();
                                if (!extTemp.IsHighResolution) extTemp.ConvertToHighResolution();
                                AssignSegmentSafely(extTemp, ext.SegmentVolume);
                                prvSeg = SafeMargin(oarTemp.SegmentVolume, marginMm);
                            }

                            var prvId = BuildId("PRV_", req.OarId);
                            SegmentVolume cappedSeg;
                            try { cappedSeg = prvSeg.And(extTemp.SegmentVolume); }
                            catch
                            {
                                cappedSeg = SafeBoolean(_ss, prvSeg, extTemp.SegmentVolume, BoolOp.And,
                                                    oarTemp, extTemp, prvId, _fb,
                                                    $"PRV_{req.OarId}_CapExt", tg);
                            }

                            if (cappedSeg != null)
                            {
                                var prv = GetOrCreate(_ss, "CONTROL", prvId);
                                if (oarTemp.IsHighResolution && !prv.IsHighResolution)
                                    prv.ConvertToHighResolution();

                                if (AssignSegmentSafely(prv, cappedSeg))
                                {
                                    prv.Color = Color.FromRgb(255, 165, 0);
                                    LogCreated(prvId);
                                }
                                else
                                {
                                    _ss.RemoveStructure(prv);
                                    _progress.AppendLine($"  SKIP: {prvId} (Empty volume)");
                                }
                            }
                        }
                    }
                    catch (Exception exPrv)
                    {
                        _progress.AppendLine($"  FAIL: PRV_{req.OarId} -> {exPrv.Message}");
                    }
                }
            }

            // ------------------------------------------------------------------
            // STEP 9: zRing_{dose}_1 / _2
            // NOTE: _zOptDoseSum is keyed by dose only; if two OptKeys share the same
            //       dose (different suffix), only the last written entry is used here.
            //       See KNOWN LIMITATIONS at the top of the file.
            // ------------------------------------------------------------------
            private void Step9_Rings(
                List<double> doseLevels, Structure ext, SegmentVolume extMinus3)
            {
                LogSection("9) zRing_{dose}_X");
                var rings = new Dictionary<double, Structure>();
                var rings2 = new Dictionary<double, Structure>();

                foreach (var d in doseLevels)
                {
                    string doseStr = d.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    if (!_zOptDoseSum.TryGetValue(d, out var optSt)) continue;

                    try
                    {
                        using (var tg = new TempGuard(_ss))
                        {
                            var optSegD = CloneSegViaTempTracked(_ss, optSt.SegmentVolume, "zTmpOptRing", tg);
                            var baseSeg = SafeMargin(optSegD, +2.0);

                            var ringOuter10 = SafeMargin(baseSeg, +RING_OUTER_EXPAND_MM);
                            var ringSeg = SafeBoolean(_ss, ringOuter10, baseSeg, BoolOp.Sub,
                                                null, optSt, $"zRing_{doseStr}_1", _fb,
                                                $"Ring1_{doseStr}_BasePlus10MinusBase", tg);

                            var ringOuter20 = SafeMargin(baseSeg, +(2.0 * RING_OUTER_EXPAND_MM));
                            var ring2Seg = SafeBoolean(_ss, ringOuter20, ringOuter10, BoolOp.Sub,
                                                null, optSt, $"zRing_{doseStr}_2", _fb,
                                                $"Ring2_{doseStr}_BasePlus20MinusBasePlus10", tg);

                            foreach (var hd in doseLevels.Where(x => x > d))
                            {
                                string hDoseStr = hd.ToString(System.Globalization.CultureInfo.InvariantCulture);

                                if (_zOptDoseSum.TryGetValue(hd, out var higherOpt))
                                {
                                    var hBase = SafeMargin(CloneSegViaTempTracked(
                                        _ss, higherOpt.SegmentVolume, "zTmpOptRingH", tg), +2.0);
                                    ringSeg = SafeBoolean(_ss, ringSeg, hBase, BoolOp.Sub, null, higherOpt, $"zRing_{doseStr}_1", _fb, $"Ring1_{doseStr}_SubOptBase_{hDoseStr}", tg);
                                    ring2Seg = SafeBoolean(_ss, ring2Seg, hBase, BoolOp.Sub, null, higherOpt, $"zRing_{doseStr}_2", _fb, $"Ring2_{doseStr}_SubOptBase_{hDoseStr}", tg);
                                }
                                if (rings.TryGetValue(hd, out var hRing1))
                                {
                                    var hRing1Seg = CloneSegViaTempTracked(_ss, hRing1.SegmentVolume, "zTmpRing1H", tg);
                                    ringSeg = SafeBoolean(_ss, ringSeg, hRing1Seg, BoolOp.Sub, null, hRing1, $"zRing_{doseStr}_1", _fb, $"Ring1_{doseStr}_SubRing1_{hDoseStr}", tg);
                                    ring2Seg = SafeBoolean(_ss, ring2Seg, hRing1Seg, BoolOp.Sub, null, hRing1, $"zRing_{doseStr}_2", _fb, $"Ring2_{doseStr}_SubRing1_{hDoseStr}", tg);
                                }
                                if (rings2.TryGetValue(hd, out var hRing2))
                                {
                                    var hRing2Seg = CloneSegViaTempTracked(_ss, hRing2.SegmentVolume, "zTmpRing2H", tg);
                                    ringSeg = SafeBoolean(_ss, ringSeg, hRing2Seg, BoolOp.Sub, null, hRing2, $"zRing_{doseStr}_1", _fb, $"Ring1_{doseStr}_SubRing2_{hDoseStr}", tg);
                                    ring2Seg = SafeBoolean(_ss, ring2Seg, hRing2Seg, BoolOp.Sub, null, hRing2, $"zRing_{doseStr}_2", _fb, $"Ring2_{doseStr}_SubRing2_{hDoseStr}", tg);
                                }
                            }

                            ringSeg = SafeBoolean(_ss, ringSeg, extMinus3, BoolOp.And, null, ext, $"zRing_{doseStr}_1", _fb, $"Ring1_{doseStr}_CapExtMinus3", tg);
                            ring2Seg = SafeBoolean(_ss, ring2Seg, extMinus3, BoolOp.And, null, ext, $"zRing_{doseStr}_2", _fb, $"Ring2_{doseStr}_CapExtMinus3", tg);

                            if (ringSeg != null)
                            {
                                var ring = GetOrCreate(_ss, "CONTROL", $"zRing_{doseStr}_1");
                                if (AssignSegmentSafely(ring, ringSeg))
                                {
                                    ring.Color = Colors.Magenta;
                                    rings[d] = ring;
                                    LogCreated($"zRing_{doseStr}_1");
                                }
                                else
                                {
                                    _ss.RemoveStructure(ring);
                                    _progress.AppendLine($"  SKIP: zRing_{doseStr}_1 (Empty volume)");
                                }
                            }

                            if (ring2Seg != null)
                            {
                                var ring2 = GetOrCreate(_ss, "CONTROL", $"zRing_{doseStr}_2");
                                if (AssignSegmentSafely(ring2, ring2Seg))
                                {
                                    ring2.Color = Colors.Magenta;
                                    rings2[d] = ring2;
                                    LogCreated($"zRing_{doseStr}_2");
                                }
                                else
                                {
                                    _ss.RemoveStructure(ring2);
                                    _progress.AppendLine($"  SKIP: zRing_{doseStr}_2 (Empty volume)");
                                }
                            }
                        }
                    }
                    catch (Exception exRing)
                    {
                        _progress.AppendLine($"  FAIL: zRing_{doseStr}_1 / _2 -> {exRing.Message}");
                    }
                }
            }
        }

        // =======================
        // HELPER METHODS (static)
        // =======================

        private static Structure FindIpsilateralLung(StructureSet ss, bool isLeft)
        {
            var lungs = ss.Structures
                .Where(s => s != null && !s.IsEmpty &&
                       (ContainsToken(s.Id, "Lung") ||
                        string.Equals(s.DicomType, "LUNG", StringComparison.OrdinalIgnoreCase)))
                .ToList();
            if (lungs.Count == 0) return null;

            string[] sideTokens = isLeft
                ? new[] { "_L", "Left", "-L", " L" }
                : new[] { "_R", "Right", "-R", " R" };

            foreach (var lung in lungs)
            {
                if (sideTokens.Any(t =>
                    lung.Id.EndsWith(t, StringComparison.OrdinalIgnoreCase) ||
                    ContainsToken(lung.Id, t)))
                    return lung;
            }

            return lungs.Count == 1 ? lungs[0] : null;
        }

        private static SegmentVolume SafeAsymmetricMargin(
            StructureSet ss, SegmentVolume sv, double mm,
            bool isLeft, SliceRecontourFallback fb, TempGuard tg, string ctx)
        {
            if (sv == null) return null;
            if (Math.Abs(mm) < 1e-6) return sv;

            double x1 = isLeft ? 0 : mm;
            double y1 = mm;
            double z1 = 0;
            double x2 = isLeft ? mm : 0;
            double y2 = 0;
            double z2 = 0;

            var margins = new AxisAlignedMargins(
                StructureMarginGeometry.Outer, x1, y1, z1, x2, y2, z2);
            try
            {
                return sv.AsymmetricMargin(margins);
            }
            catch
            {
                fb?.MarkCreated(ctx + "_IsoFallback");
                return sv.Margin(mm);
            }
        }

        // ------------------------------------------------------------------
        // v3.0.0.34: Expansion whose direction is the body/skin surface's own
        // outward normal at the point nearest the target's tip, rather than a
        // fixed global anterior/lateral axis split.
        //
        // Method: find the point on `target` closest to any point on `body`
        // (the "tip"), take that nearest body-surface point, and estimate the
        // 2D outward normal there from its neighbouring contour points on the
        // same image slice (tangent from prev/next point, rotated 90 degrees,
        // sign corrected to point away from that slice's contour centroid).
        // That single unit direction is then decomposed into the codebase's
        // existing lateral (x1/x2) and anterior (y1) AsymmetricMargin slots,
        // scaled so their combined magnitude is exactly `mm` along the
        // computed direction - reusing the same isLeft-driven x1-vs-x2 sign
        // convention already relied on elsewhere in this file, just with a
        // direction-weighted ratio instead of an even mm/mm split.
        //
        // Falls back to SafeAsymmetricMargin's even split if a tip/normal
        // can't be determined (e.g. missing contour data).
        // ------------------------------------------------------------------
        private static SegmentVolume SafePerpendicularMargin(
            StructureSet ss, Structure target, Structure body, double mm,
            bool isLeft, SliceRecontourFallback fb, TempGuard tg, string ctx)
        {
            if (target == null || target.IsEmpty) return null;
            if (Math.Abs(mm) < 1e-6) return target.SegmentVolume;

            double lateralMag = mm;
            double anteriorMag = mm;

            if (TryComputeSkinNormalAtTip(ss, target, body, out double dirX, out double dirY))
            {
                lateralMag = Math.Abs(dirX) * mm;
                anteriorMag = Math.Abs(dirY) * mm;
            }
            else
            {
                fb?.MarkCreated(ctx + "_NormalFallback");
            }

            double x1 = isLeft ? 0 : lateralMag;
            double y1 = anteriorMag;
            double x2 = isLeft ? lateralMag : 0;

            var margins = new AxisAlignedMargins(StructureMarginGeometry.Outer, x1, y1, 0, x2, 0, 0);
            try
            {
                return target.SegmentVolume.AsymmetricMargin(margins);
            }
            catch
            {
                fb?.MarkCreated(ctx + "_IsoFallback");
                return target.SegmentVolume.Margin(mm);
            }
        }

        // Finds the point on `target` nearest to `body`'s surface, then
        // estimates the outward 2D (in-plane) normal of `body` at that
        // nearest point. Returns false if either structure has no contour
        // data to work with. All distances/directions are computed directly
        // from raw contour coordinates - only the RATIO between dirX and
        // dirY is used by the caller, so no assumption about which raw axis
        // sign means "right" vs "left" is required.
        private static bool TryComputeSkinNormalAtTip(
            StructureSet ss, Structure target, Structure body,
            out double dirX, out double dirY)
        {
            dirX = 0; dirY = 0;
            if (ss?.Image == null || target == null || target.IsEmpty ||
                body == null || body.IsEmpty)
                return false;

            int nz = ss.Image.ZSize;
            double bestDist2 = double.MaxValue;
            VVector bestSkinPoint = default(VVector);
            VVector[] bestLoop = null;
            int bestIdx = -1;

            for (int z = 0; z < nz; z++)
            {
                VVector[][] targetLoops, bodyLoops;
                try { targetLoops = target.GetContoursOnImagePlane(z); }
                catch { continue; }
                if (targetLoops == null || targetLoops.Length == 0) continue;

                try { bodyLoops = body.GetContoursOnImagePlane(z); }
                catch { continue; }
                if (bodyLoops == null || bodyLoops.Length == 0) continue;

                foreach (var tLoop in targetLoops)
                {
                    if (tLoop == null) continue;
                    foreach (var tp in tLoop)
                    {
                        foreach (var bLoop in bodyLoops)
                        {
                            if (bLoop == null || bLoop.Length < 3) continue;
                            for (int i = 0; i < bLoop.Length; i++)
                            {
                                double dx = bLoop[i].x - tp.x;
                                double dy = bLoop[i].y - tp.y;
                                double d2 = dx * dx + dy * dy;
                                if (d2 < bestDist2)
                                {
                                    bestDist2 = d2;
                                    bestSkinPoint = bLoop[i];
                                    bestLoop = bLoop;
                                    bestIdx = i;
                                }
                            }
                        }
                    }
                }
            }

            if (bestLoop == null || bestIdx < 0) return false;

            int n = bestLoop.Length;
            var prev = bestLoop[(bestIdx - 1 + n) % n];
            var next = bestLoop[(bestIdx + 1) % n];
            double tangentX = next.x - prev.x;
            double tangentY = next.y - prev.y;

            double normX = tangentY;
            double normY = -tangentX;
            double len = Math.Sqrt(normX * normX + normY * normY);
            if (len < 1e-6) return false;
            normX /= len;
            normY /= len;

            double cx = 0, cy = 0;
            foreach (var p in bestLoop) { cx += p.x; cy += p.y; }
            cx /= n;
            cy /= n;

            double toPointX = bestSkinPoint.x - cx;
            double toPointY = bestSkinPoint.y - cy;
            if (normX * toPointX + normY * toPointY < 0)
            {
                normX = -normX;
                normY = -normY;
            }

            dirX = normX;
            dirY = normY;
            return true;
        }

        private static bool AssignSegmentSafely(Structure target, SegmentVolume seg)
        {
            if (target == null || seg == null) return false;
            try
            {
                target.SegmentVolume = seg;
                return !target.IsEmpty;
            }
            catch { return false; }
        }

        private static bool IsTargetByName(Structure s)
        {
            if (s == null) return false;
            var id = s.Id ?? "";
            if (id.IndexOf("PTV", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (id.IndexOf("CTV", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (id.IndexOf("GTV", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (Regex.IsMatch(id, @"(^|_|\b)LN($|_|\b)", RegexOptions.IgnoreCase)) return true;
            return false;
        }

        private static bool ContainsToken(string id, string token)
        {
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(token)) return false;
            return id.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static Structure FindExternalFallback(StructureSet ss)
        {
            if (ss == null) return null;
            return ss.Structures.FirstOrDefault(s =>
                !s.IsEmpty &&
                (string.Equals(s.DicomType, "EXTERNAL", StringComparison.OrdinalIgnoreCase)
                 || ContainsToken(s.Id, "EXTERNAL")
                 || ContainsToken(s.Id, "BODY")));
        }

        private static SegmentVolume SafeMargin(SegmentVolume sv, double mm)
        {
            if (sv == null) return null;
            if (Math.Abs(mm) < 1e-6) return sv;
            return sv.Margin(mm);
        }

        private static void SmoothStructureByExpandContract(Structure target, double mm)
        {
            if (target?.SegmentVolume == null) return;
            target.SegmentVolume = SmoothSegByExpandContract(target.SegmentVolume, mm);
        }

        private static SegmentVolume SmoothSegByExpandContract(SegmentVolume seg, double mm)
        {
            if (seg == null || mm <= 0) return seg;
            return seg.Margin(+mm).Margin(-mm);
        }

        private static SegmentVolume CloneSegViaTempTracked(
            StructureSet ss, SegmentVolume seg, string baseId, TempGuard guard)
        {
            if (seg == null) return null;
            var id = MakeUniqueId(ss, TruncId(baseId));
            var tmp = guard.Add(ss.AddStructure("CONTROL", id));
            AssignSegmentSafely(tmp, seg);
            return tmp.SegmentVolume;
        }

        private static Structure CreateTempFromSegment(
            StructureSet ss, SegmentVolume seg, string baseId)
        {
            var tmpId = MakeUniqueId(ss, TruncId(baseId));
            var tmp = ss.AddStructure("CONTROL", tmpId);
            try { tmp.SegmentVolume = seg; }
            catch
            {
                try { tmp.ConvertToHighResolution(); tmp.SegmentVolume = seg; }
                catch { }
            }
            return tmp;
        }

        private static Structure GetOrCreate(StructureSet ss, string dicomType, string id)
        {
            var existing = ss.Structures.FirstOrDefault(s =>
                string.Equals(s.Id, id, StringComparison.OrdinalIgnoreCase));
            return existing ?? ss.AddStructure(dicomType, TruncId(id));
        }

        private static string AbbreviateName(string name, int maxLen)
        {
            if (string.IsNullOrEmpty(name)) return "";
            if (name.Length <= maxLen) return name;

            var words = new List<string>();
            foreach (Match m in Regex.Matches(name, @"[A-Z]?[a-z]+|[A-Z]+(?=[A-Z][a-z]|\b)|\d+"))
                words.Add(m.Value);

            if (words.Count == 0) return name.Substring(0, Math.Min(name.Length, maxLen));

            while (words.Sum(w => w.Length) > maxLen)
            {
                int longestIdx = -1, maxWordLen = 0;
                for (int i = 0; i < words.Count; i++)
                {
                    if (words[i].Length > maxWordLen && words[i].Length > 1)
                    {
                        maxWordLen = words[i].Length;
                        longestIdx = i;
                    }
                }
                if (longestIdx == -1) break;
                words[longestIdx] = words[longestIdx].Substring(0, words[longestIdx].Length - 1);
            }

            var result = string.Join("", words);
            return result.Length > maxLen ? result.Substring(0, maxLen) : result;
        }

        private static string TruncId(string id)
        {
            if (string.IsNullOrEmpty(id) || id.Length <= 16) return id;

            string suffix = "", core = id;
            string[] knownSuffixes = { "_Left", "_Right", "_L", "_R" };
            foreach (var s in knownSuffixes)
            {
                if (core.EndsWith(s, StringComparison.OrdinalIgnoreCase))
                {
                    suffix = core.Substring(core.Length - s.Length);
                    core = core.Substring(0, core.Length - s.Length);
                    break;
                }
            }

            if (suffix.Equals("_Left", StringComparison.OrdinalIgnoreCase)) suffix = "_L";
            if (suffix.Equals("_Right", StringComparison.OrdinalIgnoreCase)) suffix = "_R";

            int budget = 16 - suffix.Length;
            if (budget <= 0) return id.Substring(0, 16);

            return AbbreviateName(core, budget) + suffix;
        }

        private static string BuildId(string prefix, string oarName, string extraSuffix = "")
        {
            prefix = prefix ?? "";
            oarName = oarName ?? "";
            extraSuffix = extraSuffix ?? "";

            string latSuffix = "", oarCore = oarName;
            string[] knownSuffixes = { "_Left", "_Right", "_L", "_R" };
            foreach (var s in knownSuffixes)
            {
                if (oarCore.EndsWith(s, StringComparison.OrdinalIgnoreCase))
                {
                    latSuffix = oarCore.Substring(oarCore.Length - s.Length);
                    oarCore = oarCore.Substring(0, oarCore.Length - s.Length);
                    break;
                }
            }

            if (latSuffix.Equals("_Left", StringComparison.OrdinalIgnoreCase)) latSuffix = "_L";
            if (latSuffix.Equals("_Right", StringComparison.OrdinalIgnoreCase)) latSuffix = "_R";

            int budget = 16 - prefix.Length - latSuffix.Length - extraSuffix.Length;
            if (budget < 1) budget = 1;

            var result = prefix + AbbreviateName(oarCore, budget) + latSuffix + extraSuffix;
            return result.Length <= 16 ? result : result.Substring(0, 16);
        }

        private static string MakeUniqueId(StructureSet ss, string baseId)
        {
            var candidate = TruncId(baseId);
            if (!ss.Structures.Any(s => string.Equals(s.Id, candidate, StringComparison.OrdinalIgnoreCase)))
                return candidate;

            for (int i = 1; i < 10000; i++)
            {
                var sfx = i.ToString();
                var room = 16 - sfx.Length;
                if (room <= 0) break;
                var next = (baseId.Length > room ? baseId.Substring(0, room) : baseId) + sfx;
                if (!ss.Structures.Any(s => string.Equals(s.Id, next, StringComparison.OrdinalIgnoreCase)))
                    return next;
            }
            throw new Exception("Could not generate a unique structure ID.");
        }

        // =======================
        // SLICE RECONTOUR FALLBACK
        // =======================
        private sealed class SliceRecontourFallback
        {
            private readonly HashSet<string> _created =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public IEnumerable<string> CreatedStructureIdsWithFallback => _created;

            public void MarkCreated(string createdId)
            {
                if (!string.IsNullOrEmpty(createdId)) _created.Add(createdId);
            }
        }

        private static void CopyContoursPlaneByPlane(
            StructureSet ss, Structure src, Structure dst)
        {
            if (src == null || dst == null || ss?.Image == null) return;
            int nz = ss.Image.ZSize;
            for (int z = 0; z < nz; z++)
            {
                var contours = src.GetContoursOnImagePlane(z);
                if (contours == null || contours.Length == 0) continue;
                foreach (var contour in contours)
                {
                    if (contour == null || contour.Length < 3) continue;
                    dst.AddContourOnImagePlane(contour, z);
                }
            }
        }

        // ==================================================================
        // BOOLEAN OPERATIONS WITH RESOLUTION MATCHING & FALLBACK
        // ==================================================================
        private enum BoolOp { And, Or, Sub }

        private static SegmentVolume TryMatchResolutionAndOp(
            StructureSet ss, SegmentVolume a, SegmentVolume b, BoolOp op, TempGuard tg)
        {
            try
            {
                var tmpA = tg.Add(ss.AddStructure("CONTROL", MakeUniqueId(ss, "zTmpMatchA")));
                try { if (a != null) tmpA.SegmentVolume = a; }
                catch { tmpA.ConvertToHighResolution(); if (a != null) tmpA.SegmentVolume = a; }

                var tmpB = tg.Add(ss.AddStructure("CONTROL", MakeUniqueId(ss, "zTmpMatchB")));
                if (tmpA.IsHighResolution && !tmpB.IsHighResolution) tmpB.ConvertToHighResolution();

                try { if (b != null) tmpB.SegmentVolume = b; }
                catch { if (!tmpB.IsHighResolution) tmpB.ConvertToHighResolution(); if (b != null) tmpB.SegmentVolume = b; }

                switch (op)
                {
                    case BoolOp.And: return tmpA.SegmentVolume.And(tmpB.SegmentVolume);
                    case BoolOp.Or: return tmpA.SegmentVolume.Or(tmpB.SegmentVolume);
                    case BoolOp.Sub: return tmpA.SegmentVolume.Sub(tmpB.SegmentVolume);
                    default: return null;
                }
            }
            catch { return null; }
        }

        private static SegmentVolume SafeBoolean(
            StructureSet ss,
            SegmentVolume a, SegmentVolume b,
            BoolOp op,
            Structure ownerA, Structure ownerB,
            string targetCreatedId,
            SliceRecontourFallback fb, string ctx, TempGuard tg)
        {
            if (op == BoolOp.And && (a == null || b == null)) return null;
            if (op == BoolOp.Or && a == null) return b;
            if (op == BoolOp.Or && b == null) return a;
            if (op == BoolOp.Sub && a == null) return null;
            if (op == BoolOp.Sub && b == null) return a;

            try
            {
                switch (op)
                {
                    case BoolOp.And: return a.And(b);
                    case BoolOp.Or: return a.Or(b);
                    case BoolOp.Sub: return a.Sub(b);
                    default: return null;
                }
            }
            catch
            {
                var viaMatch = TryMatchResolutionAndOp(ss, a, b, op, tg);
                if (viaMatch != null) return viaMatch;

                if (!string.IsNullOrEmpty(targetCreatedId)) fb?.MarkCreated(targetCreatedId);

                Structure copyA = ownerA != null
                    ? tg.Add(ss.AddStructure("CONTROL", MakeUniqueId(ss, TruncId($"zRC_{ctx}_A"))))
                    : tg.Add(CreateTempFromSegment(ss, a, "zRC_DerA"));
                if (ownerA != null) CopyContoursPlaneByPlane(ss, ownerA, copyA);

                Structure copyB = ownerB != null
                    ? tg.Add(ss.AddStructure("CONTROL", MakeUniqueId(ss, TruncId($"zRC_{ctx}_B"))))
                    : tg.Add(CreateTempFromSegment(ss, b, "zRC_DerB"));
                if (ownerB != null) CopyContoursPlaneByPlane(ss, ownerB, copyB);

                switch (op)
                {
                    case BoolOp.And: return copyA.SegmentVolume.And(copyB.SegmentVolume);
                    case BoolOp.Or: return copyA.SegmentVolume.Or(copyB.SegmentVolume);
                    case BoolOp.Sub: return copyA.SegmentVolume.Sub(copyB.SegmentVolume);
                    default: return null;
                }
            }
        }

        private static Structure UnionManyToTemp(
            StructureSet ss, IEnumerable<Structure> sources,
            SliceRecontourFallback fb, string tmpBaseId, string ctx, TempGuard tg)
        {
            var list = sources?.Where(s => s != null && !s.IsEmpty).ToList()
                       ?? new List<Structure>();
            if (list.Count == 0) return null;

            var tmpId = MakeUniqueId(ss, TruncId(tmpBaseId));
            var acc = tg.Add(ss.AddStructure("CONTROL", tmpId));
            AssignSegmentSafely(acc, list[0].SegmentVolume);

            for (int i = 1; i < list.Count; i++)
            {
                var combined = SafeBoolean(ss, acc.SegmentVolume, list[i].SegmentVolume,
                                    BoolOp.Or, acc, list[i], null, fb, ctx, tg);
                if (combined != null) AssignSegmentSafely(acc, combined);
            }
            return acc;
        }

        // ==================================================================
        // GUI MODEL + DATA CLASSES
        // ==================================================================

        public class CropOrganRow : INotifyPropertyChanged
        {
            public string OarId { get; set; }

            private bool _isTicked;
            public bool IsTicked
            {
                get => _isTicked;
                set { _isTicked = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsTicked))); }
            }

            public CropTargetData[] Targets { get; set; }
            public event PropertyChangedEventHandler PropertyChanged;
        }

        public class CropTargetData : INotifyPropertyChanged
        {
            private string _cropMm;
            public string CropMm
            {
                get => _cropMm;
                set { _cropMm = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CropMm))); }
            }
            public double? ParsedCropMm => double.TryParse(_cropMm,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : (double?)null;

            public event PropertyChangedEventHandler PropertyChanged;
        }


        private sealed class UiModel
        {
            public List<Structure> TargetCandidates { get; }
            public List<Structure> OarCandidates { get; }
            public List<Structure> ExternalCandidates { get; }
            public Structure SelectedExternal { get; set; }
            public bool? IsLeftSided { get; set; } = null;

            // v3.0.0.30: Physical bolus flag
            public bool HasPhysicalBolus { get; set; } = false;

            // v3.0.0.32: Selected physical bolus thickness (mm), added to each
            // target's virtual bolus input when Physical Bolus is ticked.
            public double PhysicalBolusThicknessMm { get; set; } = 10.0;

            public List<TargetDoseRow> TargetDoseRows { get; } = new List<TargetDoseRow>();
            public List<OrganRow> OrganRows { get; } = new List<OrganRow>();

            public UiModel(
                List<Structure> targets,
                List<Structure> oars,
                List<Structure> externals)
            {
                TargetCandidates = targets ?? new List<Structure>();
                OarCandidates = oars ?? new List<Structure>();
                ExternalCandidates = externals ?? new List<Structure>();
                SelectedExternal = ExternalCandidates.FirstOrDefault();

                foreach (var t in TargetCandidates)
                {
                    var guessedDose = OptimisationStructureWindow.GuessDoseFromName(t.Id);
                    TargetDoseRows.Add(new TargetDoseRow
                    {
                        IsSelected = false,
                        TargetId = t.Id,
                        DoseGy = guessedDose.HasValue
                                            ? guessedDose.Value.ToString("0.###",
                                                System.Globalization.CultureInfo.InvariantCulture)
                                            : "",
                        Suffix = "",
                        BolusMm = "",
                        CreateAvoidance = false,
                        CropInfo = null
                    });
                }

                foreach (var o in OarCandidates)
                {
                    OrganRows.Add(new OrganRow
                    {
                        OarId = o.Id,
                        CreateOvl = false,
                        CreateOpt = false,
                        CreatePrv = false,
                        PrvMarginMm = DEFAULT_PRV_MARGIN_MM.ToString(
                                            "0.###", System.Globalization.CultureInfo.InvariantCulture),
                        MaxDoseGy = "",
                        IsSmallOrgan = false,
                        IsLargeOrgan = false,
                        NestedSparing = false
                    });
                }
            }
        }

        private sealed class TargetDoseRow : INotifyPropertyChanged
        {
            private bool _isSelected;
            private string _doseGy;
            private string _suffix;
            private string _bolusMm;
            private bool _createAvoidance;

            public bool IsSelected
            {
                get => _isSelected;
                set { if (_isSelected != value) { _isSelected = value; OnPC(nameof(IsSelected)); } }
            }

            public string TargetId { get; set; }

            public string DoseGy
            {
                get => _doseGy;
                set { if (_doseGy != value) { _doseGy = value; OnPC(nameof(DoseGy)); } }
            }
            public double? ParsedDoseGy => double.TryParse(_doseGy,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : (double?)null;

            public string Suffix
            {
                get => _suffix;
                set { if (_suffix != value) { _suffix = value; OnPC(nameof(Suffix)); } }
            }

            public string BolusMm
            {
                get => _bolusMm;
                set { if (_bolusMm != value) { _bolusMm = value; OnPC(nameof(BolusMm)); } }
            }
            public double? ParsedBolusMm => double.TryParse(_bolusMm,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : (double?)null;

            public bool CreateAvoidance
            {
                get => _createAvoidance;
                set { if (_createAvoidance != value) { _createAvoidance = value; OnPC(nameof(CreateAvoidance)); } }
            }

            public string CropInfo { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;
            private void OnPC(string name) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private sealed class OrganRow : INotifyPropertyChanged
        {
            private bool _createOvl, _createOpt, _createPrv;
            private string _prvMarginMm;

            // ---- RCC tab fields (v3.1.0.0+) ----
            // Shared onto the same OrganRow used by the Generic/Breast Opto tabs
            // so all three tabs bind to one row type. Unused (default) outside
            // the RCC tab.
            private string _maxDoseGy;
            private bool _isSmallOrgan;
            private bool _isLargeOrgan;
            private bool _nestedSparing;
            private string _nestedThicknessMm;

            public string OarId { get; set; }

            public bool CreateOvl
            {
                get => _createOvl;
                set { if (_createOvl != value) { _createOvl = value; OnPC(nameof(CreateOvl)); } }
            }
            public bool CreateOpt
            {
                get => _createOpt;
                set { if (_createOpt != value) { _createOpt = value; OnPC(nameof(CreateOpt)); } }
            }
            public bool CreatePrv
            {
                get => _createPrv;
                set { if (_createPrv != value) { _createPrv = value; OnPC(nameof(CreatePrv)); } }
            }
            public string PrvMarginMm
            {
                get => _prvMarginMm;
                set { if (_prvMarginMm != value) { _prvMarginMm = value; OnPC(nameof(PrvMarginMm)); } }
            }
            public double? ParsedPrvMarginMm => double.TryParse(_prvMarginMm,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : (double?)null;

            // Whether this OAR participates in the Zone A max-dose crop is now
            // driven purely by whether MaxDoseGy has a value (no separate tick).
            public string MaxDoseGy
            {
                get => _maxDoseGy;
                set { if (_maxDoseGy != value) { _maxDoseGy = value; OnPC(nameof(MaxDoseGy)); } }
            }
            public double? ParsedMaxDoseGy => double.TryParse(_maxDoseGy,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : (double?)null;

            // Small/Large select which Zone A falloff rate applies to this OAR's
            // max-dose crop (small = higher %/mm, large = lower %/mm -> bigger
            // crop for the same %Diff). Mutually exclusive - ticking one clears
            // the other, enforced here so it holds regardless of which column
            // (or the header bulk-tick) changed it.
            public bool IsSmallOrgan
            {
                get => _isSmallOrgan;
                set
                {
                    if (_isSmallOrgan == value) return;
                    _isSmallOrgan = value;
                    OnPC(nameof(IsSmallOrgan));
                    if (value && _isLargeOrgan) { _isLargeOrgan = false; OnPC(nameof(IsLargeOrgan)); }
                }
            }
            public bool IsLargeOrgan
            {
                get => _isLargeOrgan;
                set
                {
                    if (_isLargeOrgan == value) return;
                    _isLargeOrgan = value;
                    OnPC(nameof(IsLargeOrgan));
                    if (value && _isSmallOrgan) { _isSmallOrgan = false; OnPC(nameof(IsSmallOrgan)); }
                }
            }
            public bool NestedSparing
            {
                get => _nestedSparing;
                set { if (_nestedSparing != value) { _nestedSparing = value; OnPC(nameof(NestedSparing)); } }
            }

            // Shell thickness (mm) used to step the §7 nested OAR-in-PTV rings
            // outward from the OAR surface; one shell per step, stopping once a
            // shell no longer overlaps the OAR (see BuildRccNestedShells).
            public string NestedThicknessMm
            {
                get => _nestedThicknessMm;
                set { if (_nestedThicknessMm != value) { _nestedThicknessMm = value; OnPC(nameof(NestedThicknessMm)); } }
            }
            public double? ParsedNestedThicknessMm => double.TryParse(_nestedThicknessMm,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : (double?)null;

            public event PropertyChangedEventHandler PropertyChanged;
            private void OnPC(string name) =>
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // ==================================================================
        // RCC PLAN / RESULT ROWS
        // These are pure result/DTO types produced by the RCC formula engine
        // (ComputeRccPlan) in SiteTabController below. They reference the same
        // TargetDoseRow / OrganRow classes used by the Generic and Breast Opto
        // tabs, since RCC now ticks/edits targets and OARs through the same
        // grids (RCC uses TargetDoseRow.DoseGy as "Rx" and adds MaxDoseGy /
        // MaxDoseGy / NestedSparing on OrganRow above).
        // ==================================================================
        private sealed class RccPlanRow
        {
            public string Category { get; set; }
            public string Source { get; set; }
            public string Zone { get; set; }
            public double PctDiff { get; set; }
            public double CropMm { get; set; }
            public string ResultId { get; set; }
        }

        private sealed class RccMaxDoseCrop
        {
            public TargetDoseRow Target;
            public OrganRow Oar;
            public double PctDiff;
            public double CropMm;
            public string ResultId;
        }

        private sealed class RccSibShave
        {
            public TargetDoseRow High;
            public TargetDoseRow Low;
            public double PctDiff;
            public double CropMm;
            public string ResultId;
        }

        private sealed class RccRingLevel
        {
            public TargetDoseRow Target;
            public double CropMm;
        }

        private sealed class RccNestedRing
        {
            public TargetDoseRow Target;
            public OrganRow Oar;
        }

        // Single source of truth shared by the plan-preview grid and the
        // structure-creation routine, so the numbers shown are the numbers used.
        private sealed class RccPlan
        {
            public List<RccMaxDoseCrop> MaxDoseCrops = new List<RccMaxDoseCrop>();
            public List<RccSibShave> SibShaves = new List<RccSibShave>();
            public double? LowestSibRxGy;
            public double Ring1TargetDoseGy;
            public double Ring2TargetDoseGy;
            public List<RccRingLevel> Ring1Levels = new List<RccRingLevel>();
            public List<RccRingLevel> Ring2Levels = new List<RccRingLevel>();
            public List<RccNestedRing> NestedRings = new List<RccNestedRing>();
        }

        // Read-only row for the "Crop Distance Matrix" grid: one row per ticked
        // OAR, one display column per ticked target, auto-computed from
        // Rx (target) + Max Dose (OAR) so the distance driving the auto-crop
        // is always visible before it is used.
        private sealed class RccMatrixRow
        {
            public string OarId { get; set; }
            public string MaxDoseDisplay { get; set; }
            public string[] CropDisplay { get; set; }
        }

        // ==================================================================
        // MAIN WINDOW
        // Hosts three independent tabs, each backed by its own UiModel and
        // built/driven by its own SiteTabController:
        //   "Generic"     - all-site crop automation (Step1-9 pipeline, no
        //                    laterality/bolus - SiteTabController.TabKind.Generic)
        //   "Breast Opto" - the original breast/chest-wall workflow, unchanged
        //                    (laterality + virtual/physical bolus -
        //                    SiteTabController.TabKind.BreastOpto)
        //   "RCC"         - "RCC Optimization Cropping Method" cheat-sheet
        //                    engine, built with the same grid-based UI shell as
        //                    the Generic tab (SiteTabController.TabKind.Rcc)
        // Only the Generic/Breast Opto tabs close the dialog on success (via
        // NotifyConfirmed, so Execute() can run the shared StructureProcessor
        // pipeline afterwards); the RCC tab creates its structures directly
        // while the dialog stays open.
        // ==================================================================
        private sealed class OptimisationStructureWindow : Window
        {
            private readonly StructureSet _ss;

            private readonly SiteTabController _tabGeneric;
            private readonly SiteTabController _tabBreast;
            private readonly SiteTabController _tabRcc;

            public UiModel ConfirmedVm { get; private set; }

            // ----------------------------------------------------------------
            // THEME
            // ----------------------------------------------------------------
            private const string ThemeXaml = @"
<ResourceDictionary xmlns=""http://schemas.microsoft.com/winfx/2006/xaml/presentation""
                    xmlns:x=""http://schemas.microsoft.com/winfx/2006/xaml"">
    <SolidColorBrush x:Key=""BgBrush""        Color=""#181B21"" />
    <SolidColorBrush x:Key=""PanelBrush""     Color=""#21252D"" />
    <SolidColorBrush x:Key=""TextPrimary""    Color=""#DCE1E8"" />
    <SolidColorBrush x:Key=""TextSecondary""  Color=""#8B94A5"" />
    <SolidColorBrush x:Key=""AccentCyan""     Color=""#00E5FF"" />
    <SolidColorBrush x:Key=""AccentPink""     Color=""#FF69B4"" />
    <SolidColorBrush x:Key=""AccentBlue""     Color=""#42A5F5"" />
    <SolidColorBrush x:Key=""CheckOrange""    Color=""#FF8C00"" />
    <SolidColorBrush x:Key=""BorderBrush""    Color=""#2E3440"" />
    <SolidColorBrush x:Key=""HoverBrush""     Color=""#2A303C"" />
    <SolidColorBrush x:Key=""WarnYellow""     Color=""#FFD54F"" />

    <Style TargetType=""TextBlock"">
        <Setter Property=""Foreground"" Value=""{StaticResource TextPrimary}"" />
        <Setter Property=""FontFamily"" Value=""Segoe UI"" />
    </Style>

    <Style TargetType=""Button"">
        <Setter Property=""Background""   Value=""{StaticResource PanelBrush}"" />
        <Setter Property=""Foreground""   Value=""{StaticResource AccentCyan}"" />
        <Setter Property=""BorderBrush""  Value=""{StaticResource BorderBrush}"" />
        <Setter Property=""BorderThickness"" Value=""1"" />
        <Setter Property=""Cursor""       Value=""Hand"" />
        <Setter Property=""Template"">
            <Setter.Value>
                <ControlTemplate TargetType=""Button"">
                    <Border Background=""{TemplateBinding Background}""
                            BorderBrush=""{TemplateBinding BorderBrush}""
                            BorderThickness=""{TemplateBinding BorderThickness}""
                            CornerRadius=""4"">
                        <ContentPresenter HorizontalAlignment=""Center""
                                          VerticalAlignment=""Center""
                                          Margin=""{TemplateBinding Padding}"" />
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
        <Style.Triggers>
            <Trigger Property=""IsMouseOver"" Value=""True"">
                <Setter Property=""Background"" Value=""{StaticResource HoverBrush}"" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style x:Key=""PrimaryButton"" TargetType=""Button"" BasedOn=""{StaticResource {x:Type Button}}"">
        <Setter Property=""Background""    Value=""{StaticResource AccentCyan}"" />
        <Setter Property=""Foreground""    Value=""#1B1E2B"" />
        <Setter Property=""BorderThickness"" Value=""0"" />
        <Setter Property=""FontWeight""    Value=""Bold"" />
        <Style.Triggers>
            <Trigger Property=""IsMouseOver"" Value=""True"">
                <Setter Property=""Opacity"" Value=""0.85"" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style x:Key=""WarningButton"" TargetType=""Button"" BasedOn=""{StaticResource {x:Type Button}}"">
        <Setter Property=""Background""    Value=""{StaticResource WarnYellow}"" />
        <Setter Property=""Foreground""    Value=""#1B1E2B"" />
        <Setter Property=""BorderThickness"" Value=""0"" />
        <Setter Property=""FontWeight""    Value=""Bold"" />
        <Style.Triggers>
            <Trigger Property=""IsMouseOver"" Value=""True"">
                <Setter Property=""Opacity"" Value=""0.85"" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style TargetType=""ComboBox"">
        <Setter Property=""Background""    Value=""{StaticResource PanelBrush}"" />
        <Setter Property=""Foreground""    Value=""#000000"" />
        <Setter Property=""BorderBrush""   Value=""{StaticResource BorderBrush}"" />
        <Setter Property=""Padding""       Value=""6,4"" />
        <Setter Property=""FontWeight""    Value=""SemiBold"" />
    </Style>

    <Style TargetType=""DataGrid"">
        <Setter Property=""Background""           Value=""Transparent"" />
        <Setter Property=""RowBackground""        Value=""Transparent"" />
        <Setter Property=""AlternatingRowBackground"" Value=""#1A1F29"" />
        <Setter Property=""BorderThickness""      Value=""0"" />
        <Setter Property=""HeadersVisibility""    Value=""Column"" />
        <Setter Property=""GridLinesVisibility""  Value=""Horizontal"" />
        <Setter Property=""HorizontalGridLinesBrush"" Value=""{StaticResource BorderBrush}"" />
        <Setter Property=""VerticalGridLinesBrush""   Value=""Transparent"" />
        <Setter Property=""Foreground""           Value=""{StaticResource TextPrimary}"" />
    </Style>

    <Style TargetType=""DataGridColumnHeader"">
        <Setter Property=""Background""    Value=""Transparent"" />
        <Setter Property=""Foreground""    Value=""{StaticResource TextSecondary}"" />
        <Setter Property=""FontWeight""    Value=""SemiBold"" />
        <Setter Property=""Padding""       Value=""10,8"" />
        <Setter Property=""BorderThickness"" Value=""0,0,0,1"" />
        <Setter Property=""BorderBrush""   Value=""{StaticResource BorderBrush}"" />
    </Style>

    <Style TargetType=""DataGridRow"">
        <Setter Property=""BorderThickness"" Value=""0"" />
        <Style.Triggers>
            <Trigger Property=""IsSelected"" Value=""True"">
                <Setter Property=""Background"" Value=""{StaticResource HoverBrush}"" />
                <Setter Property=""Foreground"" Value=""{StaticResource AccentCyan}"" />
            </Trigger>
            <Trigger Property=""IsMouseOver"" Value=""True"">
                <Setter Property=""Background"" Value=""{StaticResource HoverBrush}"" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style TargetType=""DataGridCell"">
        <Setter Property=""BorderThickness"" Value=""0"" />
        <Setter Property=""Padding""         Value=""4"" />
        <Setter Property=""Template"">
            <Setter.Value>
                <ControlTemplate TargetType=""DataGridCell"">
                    <Border Background=""{TemplateBinding Background}""
                            BorderBrush=""{TemplateBinding BorderBrush}""
                            BorderThickness=""{TemplateBinding BorderThickness}""
                            Padding=""{TemplateBinding Padding}"">
                        <ContentPresenter VerticalAlignment=""Center"" />
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
        <Style.Triggers>
            <Trigger Property=""IsSelected"" Value=""True"">
                <Setter Property=""Background"" Value=""Transparent"" />
                <Setter Property=""Foreground"" Value=""{StaticResource AccentCyan}"" />
            </Trigger>
        </Style.Triggers>
    </Style>

    <Style TargetType=""CheckBox"">
        <Setter Property=""Foreground"" Value=""{StaticResource TextPrimary}"" />
        <Setter Property=""Cursor""     Value=""Hand"" />
        <Setter Property=""Template"">
            <Setter.Value>
                <ControlTemplate TargetType=""CheckBox"">
                    <StackPanel Orientation=""Horizontal"" VerticalAlignment=""Center"" Background=""Transparent"">
                        <Border x:Name=""Border"" Width=""16"" Height=""16"" BorderThickness=""1"" CornerRadius=""3""
                                BorderBrush=""{StaticResource BorderBrush}""
                                Background=""{StaticResource PanelBrush}""
                                VerticalAlignment=""Center"">
                            <Path x:Name=""CheckMark"" Width=""9"" Height=""7"" Stretch=""Fill"" Fill=""White""
                                  Data=""M 0,3 L 3,6 L 8,0 L 9,1 L 3,8 L -1,4 Z""
                                  Visibility=""Collapsed""
                                  HorizontalAlignment=""Center"" VerticalAlignment=""Center""/>
                        </Border>
                        <ContentPresenter Margin=""8,0,0,0"" VerticalAlignment=""Center"" />
                    </StackPanel>
                    <ControlTemplate.Triggers>
                        <Trigger Property=""IsChecked"" Value=""True"">
                            <Setter TargetName=""Border""    Property=""Background""  Value=""{StaticResource CheckOrange}"" />
                            <Setter TargetName=""Border""    Property=""BorderBrush"" Value=""{StaticResource CheckOrange}"" />
                            <Setter TargetName=""CheckMark"" Property=""Visibility""  Value=""Visible"" />
                        </Trigger>
                        <Trigger Property=""IsMouseOver"" Value=""True"">
                            <Setter TargetName=""Border"" Property=""BorderBrush"" Value=""{StaticResource AccentCyan}"" />
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <Style TargetType=""RadioButton"">
        <Setter Property=""Cursor"" Value=""Hand"" />
        <Setter Property=""Template"">
            <Setter.Value>
                <ControlTemplate TargetType=""RadioButton"">
                    <StackPanel Orientation=""Horizontal"" VerticalAlignment=""Center"">
                        <Border x:Name=""Border"" Width=""18"" Height=""18"" CornerRadius=""9"" BorderThickness=""1""
                                BorderBrush=""{StaticResource BorderBrush}""
                                Background=""{StaticResource PanelBrush}"">
                            <Ellipse x:Name=""CheckMark"" Width=""10"" Height=""10""
                                     Fill=""{StaticResource CheckOrange}"" Visibility=""Collapsed""/>
                        </Border>
                        <ContentPresenter Margin=""8,0,0,0"" VerticalAlignment=""Center""/>
                    </StackPanel>
                    <ControlTemplate.Triggers>
                        <Trigger Property=""IsChecked"" Value=""True"">
                            <Setter TargetName=""Border""    Property=""BorderBrush"" Value=""{StaticResource CheckOrange}"" />
                            <Setter TargetName=""CheckMark"" Property=""Visibility""  Value=""Visible"" />
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>

    <Style TargetType=""TabControl"">
        <Setter Property=""Background"" Value=""Transparent"" />
        <Setter Property=""BorderThickness"" Value=""0"" />
    </Style>

    <Style TargetType=""TabItem"">
        <Setter Property=""Foreground"" Value=""{StaticResource TextSecondary}"" />
        <Setter Property=""FontWeight"" Value=""SemiBold"" />
        <Setter Property=""Padding"" Value=""18,10"" />
        <Setter Property=""Cursor"" Value=""Hand"" />
        <Setter Property=""Template"">
            <Setter.Value>
                <ControlTemplate TargetType=""TabItem"">
                    <Border x:Name=""Bd"" Background=""{StaticResource PanelBrush}""
                            BorderBrush=""{StaticResource BorderBrush}"" BorderThickness=""1,1,1,0""
                            CornerRadius=""6,6,0,0"" Margin=""0,0,4,0"" Padding=""{TemplateBinding Padding}"">
                        <ContentPresenter x:Name=""Cp"" ContentSource=""Header""
                                          HorizontalAlignment=""Center"" VerticalAlignment=""Center"" />
                    </Border>
                    <ControlTemplate.Triggers>
                        <Trigger Property=""IsSelected"" Value=""True"">
                            <Setter TargetName=""Bd"" Property=""Background"" Value=""{StaticResource BgBrush}"" />
                            <Setter TargetName=""Bd"" Property=""BorderBrush"" Value=""{StaticResource AccentCyan}"" />
                            <Setter Property=""Foreground"" Value=""{StaticResource AccentCyan}"" />
                        </Trigger>
                    </ControlTemplate.Triggers>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
</ResourceDictionary>
";

            public OptimisationStructureWindow(UiModel vmGeneric, UiModel vmBreast, UiModel vmRcc, StructureSet ss)
            {
                _ss = ss ?? throw new ArgumentNullException(nameof(ss));
                if (vmGeneric == null) throw new ArgumentNullException(nameof(vmGeneric));
                if (vmBreast == null) throw new ArgumentNullException(nameof(vmBreast));
                if (vmRcc == null) throw new ArgumentNullException(nameof(vmRcc));

                Title = "Generic Crop Structure Generator - v4.9.0.0";
                Width = 1250;
                Height = 800;
                MinWidth = 1000;
                MinHeight = 600;
                ResizeMode = ResizeMode.CanResize;
                WindowStartupLocation = WindowStartupLocation.CenterScreen;

                try
                {
                    var dict = (ResourceDictionary)XamlReader.Parse(ThemeXaml);
                    Resources.MergedDictionaries.Add(dict);
                    Background = (Brush)FindResource("BgBrush");
                }
                catch { }

                var inputTextBlockStyle = new Style(typeof(TextBlock));
                inputTextBlockStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty, new DynamicResourceExtension("AccentCyan")));
                inputTextBlockStyle.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.Bold));
                inputTextBlockStyle.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
                inputTextBlockStyle.Setters.Add(new Setter(TextBlock.PaddingProperty, new Thickness(4, 0, 4, 0)));

                var inputTextBoxStyle = new Style(typeof(TextBox));
                inputTextBoxStyle.Setters.Add(new Setter(TextBox.ForegroundProperty, new DynamicResourceExtension("AccentCyan")));
                inputTextBoxStyle.Setters.Add(new Setter(TextBox.BackgroundProperty, Brushes.Transparent));
                inputTextBoxStyle.Setters.Add(new Setter(TextBox.FontWeightProperty, FontWeights.Bold));
                inputTextBoxStyle.Setters.Add(new Setter(TextBox.VerticalContentAlignmentProperty, VerticalAlignment.Center));
                inputTextBoxStyle.Setters.Add(new Setter(TextBox.BorderBrushProperty, new DynamicResourceExtension("AccentCyan")));
                inputTextBoxStyle.Setters.Add(new Setter(TextBox.BorderThicknessProperty, new Thickness(0, 0, 0, 1)));
                inputTextBoxStyle.Setters.Add(new EventSetter(FrameworkElement.LoadedEvent,
                    new RoutedEventHandler((s, e) =>
                    {
                        if (s is TextBox tb) { tb.Focus(); tb.SelectAll(); }
                    })));

                Style singleClickCellStyle = new Style(typeof(DataGridCell),
                    (Style)FindResource(typeof(DataGridCell)));
                singleClickCellStyle.Setters.Add(new EventSetter(
                    UIElement.PreviewMouseLeftButtonDownEvent,
                    new System.Windows.Input.MouseButtonEventHandler((s, e) =>
                    {
                        if (s is DataGridCell cell && !cell.IsEditing && !cell.IsReadOnly &&
                            cell.Column is DataGridTextColumn)
                        {
                            if (!cell.IsFocused) cell.Focus();
                            DependencyObject parent = VisualTreeHelper.GetParent(cell);
                            while (parent != null && !(parent is DataGrid))
                                parent = VisualTreeHelper.GetParent(parent);
                            if (parent is DataGrid dg) dg.BeginEdit(e);
                        }
                    })));

                _tabGeneric = new SiteTabController(this, _ss, vmGeneric, singleClickCellStyle,
                    inputTextBlockStyle, inputTextBoxStyle, SiteTabController.TabKind.Generic,
                    "Generic crop automation (all sites). Based on Breast Gen v3.1.0.0 by Joshua Southwell, " +
                    "Medical Physicist, Australian Volunteer at LMH  ·  Modified by Mr. Ratha San, LMH, CMP.");

                _tabBreast = new SiteTabController(this, _ss, vmBreast, singleClickCellStyle,
                    inputTextBlockStyle, inputTextBoxStyle, SiteTabController.TabKind.BreastOpto,
                    "Author: Joshua Southwell, Medical Physicist, Australian Volunteer at LMH | Breast Gen v3.0.0.30  ·  Modified by Mr. Ratha San, LMH, CMP.");

                _tabRcc = new SiteTabController(this, _ss, vmRcc, singleClickCellStyle,
                    inputTextBlockStyle, inputTextBoxStyle, SiteTabController.TabKind.Rcc,
                    "RCC Optimization Cropping Method v1.0  ·  falloff-zone auto-crop from Rx + OAR Max Dose.");

                Content = BuildRootTabs();
            }

            private UIElement BuildRootTabs()
            {
                var tabs = new TabControl();

                var tabGeneric = new TabItem { Header = "Generic", Content = _tabGeneric.RootElement };
                var tabBreast = new TabItem { Header = "Breast Opto", Content = _tabBreast.RootElement };
                var tabRcc = new TabItem { Header = "RCC", Content = _tabRcc.RootElement };

                tabs.Items.Add(tabGeneric);
                tabs.Items.Add(tabBreast);
                tabs.Items.Add(tabRcc);
                tabs.SelectedIndex = 0;

                return tabs;
            }

            // Called by a SiteTabController (Generic or Breast Opto kind) once
            // CommitSelections() succeeds: closes the dialog so Execute() can run
            // the shared StructureProcessor pipeline against the confirmed vm.
            internal void NotifyConfirmed(UiModel vm)
            {
                ConfirmedVm = vm;
                DialogResult = true;
                Close();
            }

            // Closes without confirming (Cancel buttons on every tab), and also
            // used by RCC's Generate Structure button on a clean success - RCC
            // builds its structures directly rather than through NotifyConfirmed,
            // so "close the dialog" and "cancel the dialog" are the same call here.
            internal void CancelDialog()
            {
                DialogResult = false;
                Close();
            }

            private Border CreateCard(UIElement content) => new Border
            {
                Background = (Brush)FindResource("PanelBrush"),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(15),
                Margin = new Thickness(5),
                Child = content
            };

            private static bool IsIpsilateralPriority(OrganRow row, bool isLeft)
            {
                bool isPrio = PRIORITY_OARS.Any(p => ContainsToken(row.OarId, p));
                if (!isPrio) return false;

                if (ContainsToken(row.OarId, "brachial") || ContainsToken(row.OarId, "lung"))
                {
                    bool hasLeft = ContainsToken(row.OarId, "left") ||
                                    row.OarId.EndsWith("_l", StringComparison.OrdinalIgnoreCase) ||
                                    row.OarId.EndsWith("-l", StringComparison.OrdinalIgnoreCase) ||
                                    row.OarId.EndsWith(" l", StringComparison.OrdinalIgnoreCase);
                    bool hasRight = ContainsToken(row.OarId, "right") ||
                                    row.OarId.EndsWith("_r", StringComparison.OrdinalIgnoreCase) ||
                                    row.OarId.EndsWith("-r", StringComparison.OrdinalIgnoreCase) ||
                                    row.OarId.EndsWith(" r", StringComparison.OrdinalIgnoreCase);

                    if (isLeft && hasRight) return false;
                    if (!isLeft && hasLeft) return false;
                }

                return true;
            }

            private StackPanel MakeHeaderCheckbox(string label, Action<bool> onToggle)
            {
                var panel = new StackPanel { Orientation = Orientation.Horizontal };
                var cb = new CheckBox
                {
                    IsChecked = false,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 4, 0),
                    ToolTip = $"Tick/untick {label} for all rows"
                };
                cb.Checked += (s, e) => onToggle(true);
                cb.Unchecked += (s, e) => onToggle(false);
                panel.Children.Add(cb);
                panel.Children.Add(new TextBlock
                {
                    Text = label,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = FontWeights.Bold,
                    Foreground = (Brush)FindResource("TextSecondary")
                });
                return panel;
            }

            private DataGridTemplateColumn MakeSingleClickCheckColumn(
                object header, string bindingPath, double width)
            {
                var cellFactory = new FrameworkElementFactory(typeof(CheckBox));
                var binding = new Binding(bindingPath)
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                };
                cellFactory.SetBinding(CheckBox.IsCheckedProperty, binding);
                cellFactory.SetValue(CheckBox.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                cellFactory.SetValue(CheckBox.VerticalAlignmentProperty, VerticalAlignment.Center);

                return new DataGridTemplateColumn
                {
                    Header = header,
                    CellTemplate = new DataTemplate { VisualTree = cellFactory },
                    Width = new DataGridLength(width),
                    CanUserSort = false
                };
            }

            public static double? GuessDoseFromName(string id)
            {
                if (string.IsNullOrEmpty(id)) return null;
                var m = Regex.Match(id, @"(?<!\d)(\d{2,3}(\.\d{1,3})?)(?!\d)");
                if (m.Success && double.TryParse(m.Value,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var n))
                    return n;
                return null;
            }

            public static string GuessSuffixFromName(string id)
            {
                if (string.IsNullOrEmpty(id)) return "";
                string s = Regex.Replace(id, @"^(PTV|CTV|GTV|LN|ITV)", "", RegexOptions.IgnoreCase);
                s = Regex.Replace(s, @"(?<!\d)(\d{2,4}(\.\d{1,3})?)(?!\d)", "");
                s = Regex.Replace(s, @"[\s\-_]+", "_");
                return s.Trim('_');
            }

            private static bool IsSafeIdFragment(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return false;
                foreach (var c in s.Trim())
                    if (!(char.IsLetterOrDigit(c) || c == '_' || c == '-' || c == '.'))
                        return false;
                return true;
            }

            // ==============================================================
            // SITE TAB CONTROLLER
            //
            // Owns one tab's grids/state/logic. "Generic" and "Breast Opto"
            // share the exact same Step1-9 pipeline (StructureProcessor,
            // unchanged) and the exact same grid-based UI shell built by
            // Build() below; Breast Opto additionally shows the laterality +
            // physical/virtual bolus controls that Generic hides. "RCC" reuses
            // the identical shell (same top bar / left targets grid / right
            // OAR grid / bottom bar layout) but swaps in the RCC falloff-zone
            // formula engine (ported from the original RCC tab, now reading
            // Rx from TargetDoseRow.DoseGy and Max Dose / Nested-sparing from
            // the new OrganRow properties) and creates its structures directly
            // rather than through StructureProcessor.
            // ==============================================================
            private sealed class SiteTabController
            {
                public enum TabKind { Generic, BreastOpto, Rcc }

                private readonly OptimisationStructureWindow _owner;
                private readonly StructureSet _ss;
                private readonly UiModel _vm;
                private readonly TabKind _kind;
                private readonly string _footerText;
                private readonly Style _singleClickCellStyle;
                private readonly Style _inputTextBlockStyle;
                private readonly Style _inputTextBoxStyle;

                private bool IsBreast => _kind == TabKind.BreastOpto;
                private bool IsRcc => _kind == TabKind.Rcc;

                private DataGrid _dgTargets, _dgOrgans, _dgCrop;
                private ComboBox _cbExternal, _cbTargetFilter, _cbMode, _cbPhysicalThickness;

                // Only the Targets columns that SwitchMode toggles Visibility on
                // (crop-from-OAR sub-mode) need to be kept as fields; the Organs
                // grid columns are never read again after being built, so they
                // stay as local variables in BuildRightPanel instead.
                private DataGridColumn _colDose, _colSuffix, _colBolus, _colAvoid;

                private Button _btnCrop, _btnDone, _btnCreate;
                private Button _btnAutoCrop;
                private bool _inCropMode;   // Generic / Breast Opto: crop-from-OAR sub-mode

                private TextBlock _txtStats;
                private int _peakProjectedStructures;

                private List<CropOrganRow> _cropRows;
                private List<TargetDoseRow> _lastSelectedTargets;

                // ---- RCC-only state ----
                private RccPlan _rccPlan;
                private DataGrid _dgRccMatrix, _dgRccPlan;
                private TextBox _txtRccZoneASmall, _txtRccZoneALarge, _txtRccZoneB, _txtRccZoneC;

                public UIElement RootElement { get; }

                public SiteTabController(
                    OptimisationStructureWindow owner, StructureSet ss, UiModel vm,
                    Style singleClickCellStyle, Style inputTextBlockStyle, Style inputTextBoxStyle,
                    TabKind kind, string footerText)
                {
                    _owner = owner;
                    _ss = ss;
                    _vm = vm;
                    _singleClickCellStyle = singleClickCellStyle;
                    _inputTextBlockStyle = inputTextBlockStyle;
                    _inputTextBoxStyle = inputTextBoxStyle;
                    _kind = kind;
                    _footerText = footerText;

                    Action<object, PropertyChangedEventArgs> rowPropertyChanged = (s, e) =>
                    {
                        if (IsRcc) { RefreshRccPlan(); return; }

                        if (e.PropertyName == nameof(TargetDoseRow.IsSelected))
                        {
                            var targetRow = s as TargetDoseRow;
                            if (targetRow != null && targetRow.IsSelected &&
                                string.IsNullOrWhiteSpace(targetRow.Suffix))
                            {
                                targetRow.Suffix = GuessSuffixFromName(targetRow.TargetId);
                            }

                            if (_inCropMode)
                                BuildCropGrid(_vm.TargetDoseRows.Where(r => r.IsSelected).ToList());
                        }
                        UpdateStructureCount();
                    };

                    foreach (var row in _vm.TargetDoseRows)
                        row.PropertyChanged += new PropertyChangedEventHandler(rowPropertyChanged);
                    foreach (var row in _vm.OrganRows)
                        row.PropertyChanged += new PropertyChangedEventHandler(rowPropertyChanged);

                    RootElement = Build();

                    if (IsRcc) RefreshRccPlan();
                    else UpdateStructureCount();
                }

                // ----------------------------------------------------------------
                // SHARED SHELL: top bar / left targets card / right OAR card /
                // bottom bar. Kind-specific parts are built by the Build*Section
                // helpers below so all three tabs stay visually identical except
                // where the task at hand genuinely differs.
                // ----------------------------------------------------------------
                private UIElement Build()
                {
                    var root = new Grid { Margin = new Thickness(15) };
                    root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                    root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                    root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                    var topCard = _owner.CreateCard(BuildTopBar());
                    Grid.SetRow(topCard, 0);
                    root.Children.Add(topCard);

                    // RCC: targets column reduced 20% (1.15 -> 0.92), the removed
                    // width handed to the OAR/matrix column (0.8 -> 1.03) since RCC's
                    // right panel carries the crop matrix + advanced plan on top of
                    // the OAR grid. Generic / Breast Opto keep the original 1.15/0.8 split.
                    double leftStar = IsRcc ? 0.92 : 1.15;
                    double rightStar = IsRcc ? 1.03 : 0.8;

                    var mid = new Grid();
                    mid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(leftStar, GridUnitType.Star) });
                    mid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(rightStar, GridUnitType.Star) });

                    var leftCard = _owner.CreateCard(BuildTargetsPanel());
                    Grid.SetColumn(leftCard, 0);
                    mid.Children.Add(leftCard);

                    var rightCard = _owner.CreateCard(BuildRightPanel());
                    Grid.SetColumn(rightCard, 1);
                    mid.Children.Add(rightCard);

                    Grid.SetRow(mid, 1);
                    root.Children.Add(mid);

                    var bottomCard = _owner.CreateCard(BuildBottomBar());
                    Grid.SetRow(bottomCard, 2);
                    root.Children.Add(bottomCard);

                    ApplyTargetFilter();
                    return root;
                }

                // --- TOP BAR ---
                private UIElement BuildTopBar()
                {
                    var topBar = new StackPanel { Orientation = Orientation.Horizontal };

                    topBar.Children.Add(new TextBlock
                    {
                        Text = IsRcc ? "View: " : "PTV Mode: ",
                        VerticalAlignment = VerticalAlignment.Center,
                        FontWeight = FontWeights.Bold,
                        Margin = new Thickness(0, 0, 4, 0)
                    });
                    _cbMode = new ComboBox { MinWidth = 190 };
                    if (IsRcc)
                    {
                        // View-only for RCC: Generate Structure is the single
                        // action button, always visible; this just previews what
                        // it will build (OARs list vs the Advanced plan grid).
                        _cbMode.Items.Add("OARs");
                        _cbMode.Items.Add("Advanced Plan Preview");
                    }
                    else
                    {
                        _cbMode.Items.Add("Eval PTV (standard)");
                        _cbMode.Items.Add("Crop PTV from OARs");
                    }
                    _cbMode.SelectedIndex = 0;
                    _cbMode.SelectionChanged += (s, e) => SwitchMode(_cbMode.SelectedIndex == 1);
                    topBar.Children.Add(_cbMode);

                    topBar.Children.Add(new TextBlock { Text = "    Show: ", VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 4, 0) });
                    _cbTargetFilter = new ComboBox { MinWidth = 180 };
                    _cbTargetFilter.Items.Add("PTV only");
                    _cbTargetFilter.Items.Add("All (PTV, CTV, GTV, LN)");
                    _cbTargetFilter.SelectedIndex = 0;
                    _cbTargetFilter.SelectionChanged += (s, e) => ApplyTargetFilter();
                    topBar.Children.Add(_cbTargetFilter);

                    topBar.Children.Add(new TextBlock { Text = "    External/Body: ", VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 6, 0) });
                    _cbExternal = new ComboBox { MinWidth = 200 };
                    foreach (var e in _vm.ExternalCandidates) _cbExternal.Items.Add(e.Id);
                    if (_vm.SelectedExternal != null) _cbExternal.SelectedItem = _vm.SelectedExternal.Id;
                    _cbExternal.SelectionChanged += (s, e) =>
                    {
                        var id = _cbExternal.SelectedItem as string;
                        _vm.SelectedExternal = _vm.ExternalCandidates.FirstOrDefault(x =>
                            string.Equals(x.Id, id, StringComparison.OrdinalIgnoreCase));
                    };
                    topBar.Children.Add(_cbExternal);

                    return topBar;
                }

                // RCC-only: falloff zone-rate inputs, shown below the Targets grid
                // (rather than in the top bar, so the top bar stays identical
                // across all three tabs) as their own titled section - one row,
                // one column per zone (A-small, A-large, B, C).
                // Zone A is split small/large (per-OAR choice via the Organs
                // grid's Small/Large columns); B/C are the SIB-shave/ring rates.
                private UIElement BuildRccFalloffZonePanel()
                {
                    var section = new StackPanel { Orientation = Orientation.Vertical };
                    section.Children.Add(new TextBlock
                    {
                        Text = "FALLOFF ZONE",
                        FontWeight = FontWeights.Bold,
                        Foreground = (Brush)_owner.FindResource("TextSecondary"),
                        Margin = new Thickness(0, 0, 0, 8)
                    });

                    var row = new UniformGrid { Rows = 1, Columns = 4 };
                    _txtRccZoneASmall = AddRateColumn(row, "Zone A - small organ (%/mm)", RCC_ZONE_A_SMALL_DEFAULT_PCT_PER_MM);
                    _txtRccZoneALarge = AddRateColumn(row, "Zone A - large organ (%/mm)", RCC_ZONE_A_LARGE_DEFAULT_PCT_PER_MM);
                    _txtRccZoneB = AddRateColumn(row, "Zone B - SIB / ring1 (%/mm)", RCC_ZONE_B_DEFAULT_PCT_PER_MM);
                    _txtRccZoneC = AddRateColumn(row, "Zone C - ring2 (%/mm)", RCC_ZONE_C_DEFAULT_PCT_PER_MM);
                    section.Children.Add(row);

                    return _owner.CreateCard(section);
                }

                // Laterality + Physical Bolus row, shown in the left Targets
                // panel header for the Breast Opto tab only (v3.0.0.30), exactly
                // as in the original single-tab layout. Generic and RCC never
                // build this.
                private UIElement BuildLateralityBolusPanel()
                {
                    var sidePanel = new StackPanel { Orientation = Orientation.Horizontal };

                    sidePanel.Children.Add(new TextBlock
                    {
                        Text = "Laterality: ",
                        FontWeight = FontWeights.Bold,
                        Foreground = (Brush)_owner.FindResource("AccentCyan"),
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 8, 0)
                    });

                    var rbLeft = new RadioButton
                    {
                        Content = "Left Breast",
                        Foreground = (Brush)_owner.FindResource("AccentPink"),
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 10, 0)
                    };
                    var rbRight = new RadioButton
                    {
                        Content = "Right Breast",
                        Foreground = (Brush)_owner.FindResource("AccentPink"),
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(0, 0, 20, 0)
                    };
                    rbLeft.Checked += (s, e) => { _vm.IsLeftSided = true; ReSortAndTickOrgans(); };
                    rbRight.Checked += (s, e) => { _vm.IsLeftSided = false; ReSortAndTickOrgans(); };
                    sidePanel.Children.Add(rbLeft);
                    sidePanel.Children.Add(rbRight);

                    sidePanel.Children.Add(new Border
                    {
                        Width = 1,
                        Background = (Brush)_owner.FindResource("BorderBrush"),
                        Margin = new Thickness(0, 2, 20, 2),
                        VerticalAlignment = VerticalAlignment.Stretch
                    });

                    var cbPhysicalBolus = new CheckBox
                    {
                        Content = "Physical Bolus present",
                        Foreground = (Brush)_owner.FindResource("WarnYellow"),
                        FontWeight = FontWeights.SemiBold,
                        VerticalAlignment = VerticalAlignment.Center,
                        IsChecked = false,
                        ToolTip = "Tick if a BOLUS-type structure already exists in this structure set.\n" +
                                  "The virtual bolus will be generated on top of the physical bolus surface.\n" +
                                  "Also creates Bolus_physical (copy of the bolus) and Bolus_phys_Opt " +
                                  "(its overlap with z_Virtual_PTV).\n" +
                                  "Select the physical bolus thickness at right - it's added to each " +
                                  "target's bolus input (mm) to build z_Virtual_PTV."
                    };

                    _cbPhysicalThickness = new ComboBox
                    {
                        MinWidth = 60,
                        Margin = new Thickness(6, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center,
                        IsEnabled = false,
                        ToolTip = "Physical bolus thickness (mm), added to the virtual bolus input for z_Virtual_PTV."
                    };
                    foreach (var t in PhysicalBolusThicknessOptionsMm)
                        _cbPhysicalThickness.Items.Add(t.ToString("0") + " mm");
                    _cbPhysicalThickness.SelectedIndex = 1; // default 10 mm
                    _cbPhysicalThickness.SelectionChanged += (s, e) =>
                    {
                        int idx = Math.Max(0, _cbPhysicalThickness.SelectedIndex);
                        _vm.PhysicalBolusThicknessMm = PhysicalBolusThicknessOptionsMm[idx];
                    };

                    cbPhysicalBolus.Checked += (s, e) =>
                    {
                        _vm.HasPhysicalBolus = true;
                        _cbPhysicalThickness.IsEnabled = true;
                    };
                    cbPhysicalBolus.Unchecked += (s, e) =>
                    {
                        _vm.HasPhysicalBolus = false;
                        _cbPhysicalThickness.IsEnabled = false;
                    };
                    sidePanel.Children.Add(cbPhysicalBolus);
                    sidePanel.Children.Add(_cbPhysicalThickness);

                    return sidePanel;
                }

                // Shared "tick to bulk-apply" boolean column: wires the header
                // checkbox to set `setter` on every row of `rows`, refreshes
                // `grid`, then runs `onChanged` (UpdateStructureCount /
                // RefreshRccPlan / no-op). Every plain tick column across all
                // three tabs' Targets/Organs/Crop grids goes through this
                // instead of repeating the same foreach+refresh block.
                private DataGridColumn AddBoolColumn<T>(
                    DataGrid grid, IEnumerable<T> rows, string header,
                    Action<T, bool> setter, string bindingPath, double width, Action onChanged)
                {
                    var col = _owner.MakeSingleClickCheckColumn(
                        _owner.MakeHeaderCheckbox(header, isChecked =>
                        {
                            foreach (var r in rows) setter(r, isChecked);
                            grid.Items.Refresh();
                            onChanged();
                        }),
                        bindingPath, width);
                    grid.Columns.Add(col);
                    return col;
                }

                // --- LEFT: TARGETS ---
                private UIElement BuildTargetsPanel()
                {
                    var leftPanel = new DockPanel();
                    var leftHeader = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 0, 0, 10) };

                    leftHeader.Children.Add(new TextBlock
                    {
                        Text = IsRcc ? "TARGETS  (tick to include, set Rx)" : "TARGETS  (tick to include, set params)",
                        FontWeight = FontWeights.Bold,
                        Foreground = (Brush)_owner.FindResource("TextSecondary"),
                        Margin = new Thickness(0, 0, 0, 8)
                    });

                    if (IsBreast) leftHeader.Children.Add(BuildLateralityBolusPanel());

                    DockPanel.SetDock(leftHeader, Dock.Top);
                    leftPanel.Children.Add(leftHeader);

                    _dgTargets = new DataGrid
                    {
                        AutoGenerateColumns = false,
                        CanUserAddRows = false,
                        CanUserDeleteRows = false,
                        SelectionMode = DataGridSelectionMode.Extended,
                        ItemsSource = _vm.TargetDoseRows,
                        CellStyle = _singleClickCellStyle
                    };

                    _dgTargets.Columns.Add(_owner.MakeSingleClickCheckColumn(
                        _owner.MakeHeaderCheckbox("Use", isChecked =>
                        {
                            foreach (var r in _vm.TargetDoseRows)
                            {
                                r.IsSelected = isChecked;
                                if (isChecked && string.IsNullOrWhiteSpace(r.Suffix))
                                    r.Suffix = GuessSuffixFromName(r.TargetId);
                            }
                            _dgTargets.Items.Refresh();
                            if (IsRcc) RefreshRccPlan(); else UpdateStructureCount();
                        }),
                        nameof(TargetDoseRow.IsSelected), 75));

                    var targetColStyle = new Style(typeof(TextBlock), (Style)_owner.FindResource(typeof(TextBlock)));
                    targetColStyle.Setters.Add(new Setter(TextBlock.ToolTipProperty,
                        new Binding(nameof(TargetDoseRow.CropInfo))));

                    _dgTargets.Columns.Add(new DataGridTextColumn
                    {
                        Header = "Target",
                        Binding = new Binding(nameof(TargetDoseRow.TargetId)),
                        IsReadOnly = true,
                        Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                        ElementStyle = targetColStyle
                    });

                    _colDose = new DataGridTextColumn
                    {
                        Header = IsRcc ? "Rx (Gy)" : "Dose (Gy)",
                        Binding = new Binding(nameof(TargetDoseRow.DoseGy)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                        Width = 95,
                        ElementStyle = _inputTextBlockStyle,
                        EditingElementStyle = _inputTextBoxStyle
                    };
                    _dgTargets.Columns.Add(_colDose);

                    _colSuffix = new DataGridTextColumn
                    {
                        Header = "Suffix",
                        Binding = new Binding(nameof(TargetDoseRow.Suffix)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                        Width = 85,
                        ElementStyle = _inputTextBlockStyle,
                        EditingElementStyle = _inputTextBoxStyle
                    };
                    _dgTargets.Columns.Add(_colSuffix);

                    if (IsBreast)
                    {
                        _colBolus = new DataGridTextColumn
                        {
                            Header = "Bolus (mm)",
                            Binding = new Binding(nameof(TargetDoseRow.BolusMm)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                            Width = 95,
                            ElementStyle = _inputTextBlockStyle,
                            EditingElementStyle = _inputTextBoxStyle
                        };
                        _dgTargets.Columns.Add(_colBolus);
                    }

                    if (!IsRcc)
                    {
                        _colAvoid = AddBoolColumn(_dgTargets, _vm.TargetDoseRows, "Avoid",
                            (r, v) => r.CreateAvoidance = v, nameof(TargetDoseRow.CreateAvoidance), 85, UpdateStructureCount);
                    }

                    // DockPanel stacks same-side children from the outside in, so
                    // among the Bottom-docked items the FIRST one added ends up
                    // outermost (bottom-most). Add the matrix first so it sits below
                    // the falloff-zone section, which sits below the Targets grid.
                    if (IsRcc)
                    {
                        var matrixPanel = BuildRccCropDistanceMatrixPanel();
                        DockPanel.SetDock(matrixPanel, Dock.Bottom);
                        leftPanel.Children.Add(matrixPanel);

                        var zonePanel = BuildRccFalloffZonePanel();
                        DockPanel.SetDock(zonePanel, Dock.Bottom);
                        leftPanel.Children.Add(zonePanel);
                    }

                    leftPanel.Children.Add(_dgTargets);
                    return leftPanel;
                }

                // RCC-only: the read-only "Crop Distance Matrix" (one row per OAR
                // with a Max Dose, one column per ticked target), shown below the
                // Targets section rather than in the OAR panel so the Organs grid
                // can fill the whole right-hand panel instead.
                private UIElement BuildRccCropDistanceMatrixPanel()
                {
                    var section = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 10, 0, 0) };
                    section.Children.Add(new TextBlock
                    {
                        Text = "CROP DISTANCE (auto, from Rx & Max Dose)",
                        FontWeight = FontWeights.Bold,
                        Foreground = (Brush)_owner.FindResource("TextSecondary"),
                        Margin = new Thickness(0, 0, 0, 8)
                    });

                    _dgRccMatrix = new DataGrid
                    {
                        AutoGenerateColumns = false,
                        CanUserAddRows = false,
                        IsReadOnly = true,
                        HeadersVisibility = DataGridHeadersVisibility.Column,
                        MaxHeight = 220
                    };
                    section.Children.Add(_dgRccMatrix);

                    return _owner.CreateCard(section);
                }

                // --- RIGHT: OAR/ORGANS (+ RCC advanced plan) ---
                private UIElement BuildRightPanel()
                {
                    var rightPanel = new DockPanel();
                    var rightHeader = new TextBlock
                    {
                        Text = IsRcc
                            ? "OARs  (enter Max Dose to auto-crop; tick Small/Large for falloff zone; tick Nested for §7)"
                            : "ORGANS  (tick what to create per organ)",
                        FontWeight = FontWeights.Bold,
                        Foreground = (Brush)_owner.FindResource("AccentBlue"),
                        Margin = new Thickness(0, 0, 0, 10)
                    };
                    DockPanel.SetDock(rightHeader, Dock.Top);
                    rightPanel.Children.Add(rightHeader);

                    _dgOrgans = new DataGrid
                    {
                        AutoGenerateColumns = false,
                        CanUserAddRows = false,
                        CanUserDeleteRows = false,
                        SelectionMode = DataGridSelectionMode.Extended,
                        ItemsSource = _vm.OrganRows,
                        CellStyle = _singleClickCellStyle
                    };

                    if (!IsRcc)
                    {
                        _dgOrgans.PreviewMouseLeftButtonUp += (s, e) =>
                        {
                            if (_inCropMode) return;
                            var dep = e.OriginalSource as DependencyObject;
                            while (dep != null && !(dep is DataGridCell))
                            {
                                dep = (dep is Visual || dep is System.Windows.Media.Media3D.Visual3D)
                                    ? VisualTreeHelper.GetParent(dep)
                                    : LogicalTreeHelper.GetParent(dep);
                            }
                            if (dep is DataGridCell cell && _dgOrgans.Columns.IndexOf(cell.Column) == 0)
                            {
                                if (cell.DataContext is OrganRow row)
                                {
                                    if (row.CreateOvl && row.CreateOpt)
                                    { row.CreateOvl = false; row.CreateOpt = false; row.CreatePrv = false; }
                                    else
                                    { row.CreateOvl = true; row.CreateOpt = true; }
                                    _dgOrgans.Items.Refresh();
                                    UpdateStructureCount();
                                }
                            }
                        };
                    }

                    _dgOrgans.Columns.Add(new DataGridTextColumn
                    {
                        Header = "Organ",
                        Binding = new Binding(nameof(OrganRow.OarId)),
                        IsReadOnly = true,
                        Width = new DataGridLength(140),
                        ElementStyle = (Style)_owner.FindResource(typeof(TextBlock))
                    });

                    if (IsRcc)
                    {
                        // No separate "Crop" tick - an OAR is included the moment a
                        // valid Max Dose is entered (confirmed by the number itself).
                        _dgOrgans.Columns.Add(new DataGridTextColumn
                        {
                            Header = "Max Dose (Gy)",
                            Binding = new Binding(nameof(OrganRow.MaxDoseGy)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                            Width = 110,
                            ElementStyle = _inputTextBlockStyle,
                            EditingElementStyle = _inputTextBoxStyle
                        });

                        // Small/Large pick which Zone A rate applies to this OAR
                        // (mutually exclusive - enforced on OrganRow itself).
                        AddBoolColumn(_dgOrgans, _vm.OrganRows, "Small",
                            (r, v) => r.IsSmallOrgan = v, nameof(OrganRow.IsSmallOrgan), 65, RefreshRccPlan);

                        AddBoolColumn(_dgOrgans, _vm.OrganRows, "Large",
                            (r, v) => r.IsLargeOrgan = v, nameof(OrganRow.IsLargeOrgan), 65, RefreshRccPlan);

                        AddBoolColumn(_dgOrgans, _vm.OrganRows, "Nested §7",
                            (r, v) => r.NestedSparing = v, nameof(OrganRow.NestedSparing), 90, RefreshRccPlan);

                        // Shell thickness driving how many z_oar_in_ptv_hr# rings
                        // BuildRccNestedShells steps outward (see that method).
                        _dgOrgans.Columns.Add(new DataGridTextColumn
                        {
                            Header = "Nested Thickness (mm)",
                            Binding = new Binding(nameof(OrganRow.NestedThicknessMm)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                            Width = 150,
                            ElementStyle = _inputTextBlockStyle,
                            EditingElementStyle = _inputTextBoxStyle
                        });
                    }
                    else
                    {
                        AddBoolColumn(_dgOrgans, _vm.OrganRows, "Ovl",
                            (r, v) => r.CreateOvl = v, nameof(OrganRow.CreateOvl), 75, UpdateStructureCount);

                        AddBoolColumn(_dgOrgans, _vm.OrganRows, "Opt",
                            (r, v) => r.CreateOpt = v, nameof(OrganRow.CreateOpt), 75, UpdateStructureCount);

                        AddBoolColumn(_dgOrgans, _vm.OrganRows, "PRV",
                            (r, v) => r.CreatePrv = v, nameof(OrganRow.CreatePrv), 75, UpdateStructureCount);

                        _dgOrgans.Columns.Add(new DataGridTextColumn
                        {
                            Header = "PRV Margin (mm)",
                            Binding = new Binding(nameof(OrganRow.PrvMarginMm)) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                            Width = 150,
                            ElementStyle = _inputTextBlockStyle,
                            EditingElementStyle = _inputTextBoxStyle
                        });
                    }

                    if (IsRcc)
                    {
                        // Organs and the Advanced plan preview share the same Grid
                        // cell (only one Visible at a time, toggled by SwitchMode) so
                        // whichever is showing fills all remaining right-panel space
                        // - a plain DockPanel "last child fills" rule only applies to
                        // the literal last child, not whichever sibling is visible.
                        _dgRccPlan = new DataGrid
                        {
                            AutoGenerateColumns = false,
                            CanUserAddRows = false,
                            IsReadOnly = true,
                            HeadersVisibility = DataGridHeadersVisibility.Column,
                            Visibility = Visibility.Collapsed
                        };
                        _dgRccPlan.Columns.Add(new DataGridTextColumn { Header = "Step", Binding = new Binding("Category"), Width = 150, ElementStyle = (Style)_owner.FindResource(typeof(TextBlock)) });
                        _dgRccPlan.Columns.Add(new DataGridTextColumn { Header = "Source", Binding = new Binding("Source"), Width = new DataGridLength(1, DataGridLengthUnitType.Star), ElementStyle = (Style)_owner.FindResource(typeof(TextBlock)) });
                        _dgRccPlan.Columns.Add(new DataGridTextColumn { Header = "Zone", Binding = new Binding("Zone"), Width = 45, ElementStyle = (Style)_owner.FindResource(typeof(TextBlock)) });
                        _dgRccPlan.Columns.Add(new DataGridTextColumn { Header = "%Diff", Binding = new Binding("PctDiff") { StringFormat = "0.00" }, Width = 65, ElementStyle = (Style)_owner.FindResource(typeof(TextBlock)) });
                        _dgRccPlan.Columns.Add(new DataGridTextColumn { Header = "Crop (mm)", Binding = new Binding("CropMm") { StringFormat = "0.00" }, Width = 80, ElementStyle = (Style)_owner.FindResource(typeof(TextBlock)) });
                        _dgRccPlan.Columns.Add(new DataGridTextColumn { Header = "Result ID", Binding = new Binding("ResultId"), Width = 130, ElementStyle = (Style)_owner.FindResource(typeof(TextBlock)) });

                        var toggleArea = new Grid();
                        toggleArea.Children.Add(_dgOrgans);
                        toggleArea.Children.Add(_dgRccPlan);
                        rightPanel.Children.Add(toggleArea);
                    }
                    else
                    {
                        rightPanel.Children.Add(_dgOrgans);

                        _dgCrop = new DataGrid
                        {
                            AutoGenerateColumns = false,
                            CanUserAddRows = false,
                            CanUserDeleteRows = false,
                            SelectionMode = DataGridSelectionMode.Extended,
                            Visibility = Visibility.Collapsed,
                            CellStyle = _singleClickCellStyle
                        };
                        rightPanel.Children.Add(_dgCrop);
                    }

                    return rightPanel;
                }

                // --- BOTTOM BAR ---
                private UIElement BuildBottomBar()
                {
                    var bottomGrid = new Grid();
                    bottomGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    bottomGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    var leftFooter = new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Center };
                    _txtStats = new TextBlock { Foreground = (Brush)_owner.FindResource("AccentCyan"), FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 4) };
                    leftFooter.Children.Add(_txtStats);
                    leftFooter.Children.Add(new TextBlock
                    {
                        Text = _footerText,
                        Foreground = (Brush)_owner.FindResource("TextSecondary"),
                        FontSize = 11
                    });
                    Grid.SetColumn(leftFooter, 0);
                    bottomGrid.Children.Add(leftFooter);

                    var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

                    var btnClear = new Button { Content = "Clear ticks", Margin = new Thickness(0, 0, 20, 0), Padding = new Thickness(15, 6, 15, 6) };
                    btnClear.Click += (s, e) =>
                    {
                        foreach (var r in _vm.OrganRows)
                        {
                            r.CreateOvl = false; r.CreateOpt = false; r.CreatePrv = false;
                            r.MaxDoseGy = ""; r.IsSmallOrgan = false; r.IsLargeOrgan = false; r.NestedSparing = false; r.NestedThicknessMm = "";
                        }
                        foreach (var r in _vm.TargetDoseRows) { r.IsSelected = false; r.CreateAvoidance = false; r.BolusMm = ""; }
                        if (_cropRows != null) { foreach (var r in _cropRows) r.IsTicked = false; _dgCrop?.Items.Refresh(); }
                        _dgOrgans.Items.Refresh();
                        _dgTargets.Items.Refresh();
                        if (IsRcc) RefreshRccPlan(); else UpdateStructureCount();
                    };
                    buttonPanel.Children.Add(btnClear);

                    if (IsRcc)
                    {
                        // Single unified action - single-target and multi-target
                        // (SIB) selections both go through DoRccGenerateStructure;
                        // there's no separate Advanced/SIB/Ring/Nested button anymore.
                        _btnAutoCrop = new Button { Content = "Generate Structure", Padding = new Thickness(20, 8, 20, 8), Margin = new Thickness(0, 0, 6, 0) };
                        _btnAutoCrop.SetResourceReference(FrameworkElement.StyleProperty, "PrimaryButton");
                        _btnAutoCrop.Click += (s, e) => DoRccGenerateStructure();
                        buttonPanel.Children.Add(_btnAutoCrop);
                    }
                    else
                    {
                        _btnCrop = new Button { Content = "Crop PTVs", Padding = new Thickness(20, 6, 20, 6), Margin = new Thickness(0, 0, 6, 0), Visibility = Visibility.Collapsed };
                        _btnCrop.SetResourceReference(FrameworkElement.StyleProperty, "WarningButton");
                        _btnCrop.Click += (s, e) => DoCrop();
                        buttonPanel.Children.Add(_btnCrop);

                        _btnDone = new Button { Content = "Done", Padding = new Thickness(15, 6, 15, 6), Margin = new Thickness(0, 0, 6, 0), Visibility = Visibility.Collapsed };
                        _btnDone.Click += (s, e) => { if (_cbMode != null) _cbMode.SelectedIndex = 0; };
                        buttonPanel.Children.Add(_btnDone);

                        _btnCreate = new Button { Content = "Create Structures", IsDefault = true, Padding = new Thickness(25, 8, 25, 8), Margin = new Thickness(0, 0, 6, 0) };
                        _btnCreate.SetResourceReference(FrameworkElement.StyleProperty, "PrimaryButton");
                        _btnCreate.Click += (s, e) => { if (CommitSelections()) _owner.NotifyConfirmed(_vm); };
                        buttonPanel.Children.Add(_btnCreate);
                    }

                    var btnCancel = new Button { Content = "Cancel", Padding = new Thickness(15, 6, 15, 6) };
                    btnCancel.Click += (s, e) => _owner.CancelDialog();
                    buttonPanel.Children.Add(btnCancel);

                    Grid.SetColumn(buttonPanel, 1);
                    bottomGrid.Children.Add(buttonPanel);

                    return bottomGrid;
                }

                // ----------------------------------------------------------------
                // GENERIC / BREAST OPTO: crop-from-OAR sub-mode + commit/validate
                // ----------------------------------------------------------------
                private void SwitchMode(bool secondMode)
                {
                    if (IsRcc)
                    {
                        // View-only toggle: Organs and the Advanced plan preview
                        // share one Grid cell (see BuildRightPanel), so switching
                        // Visibility here is enough for whichever is shown to fill
                        // it. Generate Structure is the one action button and stays
                        // visible either way; the Crop Distance Matrix lives under
                        // the Targets section now and is visible in both views.
                        _dgOrgans.Visibility = secondMode ? Visibility.Collapsed : Visibility.Visible;
                        _dgRccPlan.Visibility = secondMode ? Visibility.Visible : Visibility.Collapsed;
                        RefreshRccPlan();
                        return;
                    }

                    if (secondMode)
                    {
                        BuildCropGrid(_vm.TargetDoseRows.Where(r => r.IsSelected).ToList());
                        _dgOrgans.Visibility = Visibility.Collapsed;
                        _dgCrop.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        if (_inCropMode)
                        {
                            foreach (var r in _vm.OrganRows) { r.CreateOvl = false; r.CreateOpt = false; }
                            _dgOrgans.Items.Refresh();
                            UpdateStructureCount();
                        }
                        _dgOrgans.Visibility = Visibility.Visible;
                        _dgCrop.Visibility = Visibility.Collapsed;
                    }

                    _inCropMode = secondMode;

                    if (_colDose != null) _colDose.Visibility = secondMode ? Visibility.Collapsed : Visibility.Visible;
                    if (_colSuffix != null) _colSuffix.Visibility = secondMode ? Visibility.Collapsed : Visibility.Visible;
                    if (_colBolus != null) _colBolus.Visibility = secondMode ? Visibility.Collapsed : Visibility.Visible;
                    if (_colAvoid != null) _colAvoid.Visibility = secondMode ? Visibility.Collapsed : Visibility.Visible;

                    _btnCrop.Visibility = secondMode ? Visibility.Visible : Visibility.Collapsed;
                    _btnDone.Visibility = secondMode ? Visibility.Visible : Visibility.Collapsed;
                    _btnCreate.Visibility = secondMode ? Visibility.Collapsed : Visibility.Visible;
                }

                private void BuildCropGrid(List<TargetDoseRow> selectedTargets)
                {
                    var oldRows = _cropRows;
                    _cropRows = new List<CropOrganRow>();

                    foreach (var oar in _vm.OrganRows)
                    {
                        var oldRow = oldRows?.FirstOrDefault(r => r.OarId == oar.OarId);

                        var row = new CropOrganRow
                        {
                            OarId = oar.OarId,
                            IsTicked = oldRow?.IsTicked ?? false,
                            Targets = new CropTargetData[selectedTargets.Count]
                        };

                        row.PropertyChanged += (s, e) => UpdateStructureCount();

                        for (int i = 0; i < selectedTargets.Count; i++)
                        {
                            double? oldDist = null;
                            if (oldRow != null && _lastSelectedTargets != null)
                            {
                                int oldIdx = _lastSelectedTargets.FindIndex(
                                    t => t.TargetId == selectedTargets[i].TargetId);
                                if (oldIdx >= 0 && oldIdx < oldRow.Targets.Length)
                                    oldDist = oldRow.Targets[oldIdx].ParsedCropMm;
                            }
                            row.Targets[i] = new CropTargetData
                            {
                                CropMm = oldDist?.ToString("0.###",
                                    System.Globalization.CultureInfo.InvariantCulture) ?? ""
                            };
                        }
                        _cropRows.Add(row);
                    }

                    _lastSelectedTargets = new List<TargetDoseRow>(selectedTargets);

                    _dgCrop.Columns.Clear();
                    _dgCrop.Columns.Add(new DataGridTextColumn
                    {
                        Header = "Organ",
                        Binding = new Binding("OarId"),
                        IsReadOnly = true,
                        Width = new DataGridLength(140),
                        ElementStyle = (Style)_owner.FindResource(typeof(TextBlock))
                    });

                    AddBoolColumn(_dgCrop, _cropRows, "Crop?", (r, v) => r.IsTicked = v, "IsTicked", 75, () => { });

                    for (int i = 0; i < selectedTargets.Count; i++)
                    {
                        string tid = selectedTargets[i].TargetId;
                        string headerName = tid.Length > 10 ? tid.Substring(0, 10) + ".." : tid;
                        _dgCrop.Columns.Add(new DataGridTextColumn
                        {
                            Header = headerName + "(mm)",
                            Binding = new Binding($"Targets[{i}].CropMm") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.LostFocus },
                            Width = 95,
                            ElementStyle = _inputTextBlockStyle,
                            EditingElementStyle = _inputTextBoxStyle
                        });
                    }

                    _dgCrop.ItemsSource = _cropRows;
                }

                private void DoCrop()
                {
                    _dgTargets.CommitEdit(DataGridEditingUnit.Cell, true);
                    _dgTargets.CommitEdit(DataGridEditingUnit.Row, true);
                    _dgCrop.CommitEdit(DataGridEditingUnit.Cell, true);
                    _dgCrop.CommitEdit(DataGridEditingUnit.Row, true);

                    var selectedTargets = _vm.TargetDoseRows.Where(r => r.IsSelected).ToList();
                    var fb = new SliceRecontourFallback();
                    var created = new List<string>();
                    var errors = new List<string>();
                    var ext = _vm.SelectedExternal;

                    if (ext == null || ext.IsEmpty)
                    {
                        MessageBox.Show(_owner, "No External/Body structure selected.",
                            "Missing External", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    if (!_cropRows.Any(r => r.IsTicked))
                    {
                        MessageBox.Show(_owner, "Please tick at least one organ to crop from.",
                            "No organs selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int ptvIndex = 0;
                    foreach (var tdr in selectedTargets)
                    {
                        var target = _ss.Structures.FirstOrDefault(s =>
                            !s.IsEmpty && string.Equals(s.Id, tdr.TargetId, StringComparison.OrdinalIgnoreCase));
                        if (target == null) { ptvIndex++; continue; }

                        var tickedOars = _cropRows.Where(r => r.IsTicked).ToList();
                        if (tickedOars.Count == 0) { ptvIndex++; continue; }

                        using (var tgUnion = new TempGuard(_ss))
                        {
                            SegmentVolume expandedOarsUnion = null;
                            Structure expandedOarsUnionSt = null;

                            foreach (var oarRow in tickedOars)
                            {
                                var oarSt = _ss.Structures.FirstOrDefault(s =>
                                    !s.IsEmpty && string.Equals(s.Id, oarRow.OarId, StringComparison.OrdinalIgnoreCase));
                                if (oarSt == null) continue;

                                try
                                {
                                    double dist = oarRow.Targets[ptvIndex].ParsedCropMm.GetValueOrDefault();
                                    var expanded = SafeMargin(oarSt.SegmentVolume, dist);
                                    var tmpSt = tgUnion.Add(CreateTempFromSegment(_ss, expanded, "zRC_ExpOar"));

                                    if (expandedOarsUnion == null)
                                    {
                                        expandedOarsUnion = expanded;
                                        expandedOarsUnionSt = tmpSt;
                                    }
                                    else
                                    {
                                        expandedOarsUnion = SafeBoolean(_ss,
                                            expandedOarsUnion, expanded, BoolOp.Or,
                                            expandedOarsUnionSt, tmpSt, null, fb, "CropOarsUnion", tgUnion);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errors.Add($"Organ {oarRow.OarId} union failed: {ex.Message}");
                                }
                            }

                            if (expandedOarsUnion != null)
                            {
                                try
                                {
                                    using (var tgCrop = new TempGuard(_ss))
                                    {
                                        var croppedSeg = SafeBoolean(_ss,
                                            target.SegmentVolume, expandedOarsUnion, BoolOp.Sub,
                                            target, expandedOarsUnionSt, null, fb,
                                            $"Crop_{tdr.TargetId}_Sub", tgCrop);
                                        croppedSeg = SafeBoolean(_ss, croppedSeg, ext.SegmentVolume, BoolOp.And,
                                            null, ext, null, fb, $"Crop_{tdr.TargetId}_CapExt", tgCrop);

                                        bool isReCrop = tdr.TargetId.EndsWith("_Crp", StringComparison.OrdinalIgnoreCase);
                                        string cropId = isReCrop ? tdr.TargetId : BuildId("z_", tdr.TargetId, "_Crp");

                                        var cropSt = GetOrCreate(_ss, "PTV", cropId);
                                        if (AssignSegmentSafely(cropSt, croppedSeg))
                                        {
                                            cropSt.Color = Color.FromRgb(255, 165, 0);
                                            var oarNames = string.Join(", ", tickedOars.Select(r =>
                                                $"{r.OarId}({r.Targets[ptvIndex].ParsedCropMm.GetValueOrDefault().ToString("0.###", System.Globalization.CultureInfo.InvariantCulture)}mm)"));
                                            var info = $"Cropped from: {oarNames}";
                                            if (isReCrop && !string.IsNullOrEmpty(tdr.CropInfo))
                                                info = tdr.CropInfo + "\n---\n" + info;

                                            created.Add(cropId);

                                            if (isReCrop)
                                            {
                                                tdr.CropInfo = info;
                                            }
                                            else
                                            {
                                                var newRow = new TargetDoseRow
                                                {
                                                    IsSelected = false,
                                                    TargetId = cropId,
                                                    DoseGy = tdr.DoseGy,
                                                    Suffix = tdr.Suffix,
                                                    BolusMm = tdr.BolusMm,
                                                    CreateAvoidance = tdr.CreateAvoidance,
                                                    CropInfo = info
                                                };
                                                newRow.PropertyChanged += (s, e) =>
                                                {
                                                    if (e.PropertyName == nameof(TargetDoseRow.IsSelected) && _inCropMode)
                                                        BuildCropGrid(_vm.TargetDoseRows.Where(r => r.IsSelected).ToList());
                                                    UpdateStructureCount();
                                                };
                                                _vm.TargetDoseRows.Add(newRow);
                                            }
                                        }
                                        else
                                        {
                                            _ss.RemoveStructure(cropSt);
                                            errors.Add($"{tdr.TargetId}: Crop resulted in an empty structure.");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errors.Add($"{tdr.TargetId}: {ex.Message}");
                                }
                            }
                        }
                        ptvIndex++;
                    }

                    var msg = new StringBuilder();
                    msg.AppendLine($"Cropped {created.Count} PTV(s):");
                    foreach (var id in created) msg.AppendLine($"  {id}");
                    if (errors.Count > 0)
                    {
                        msg.AppendLine();
                        msg.AppendLine($"Errors ({errors.Count}):");
                        foreach (var err in errors) msg.AppendLine($"  {err}");
                    }
                    msg.AppendLine();
                    msg.AppendLine("You can crop again, or click Done to proceed.");
                    MessageBox.Show(_owner, msg.ToString(), "Crop Complete", MessageBoxButton.OK, MessageBoxImage.Information);

                    _dgTargets.Items.Refresh();
                    ApplyTargetFilter();
                }

                private void ApplyTargetFilter()
                {
                    if (_dgTargets == null || _cbTargetFilter == null) return;
                    bool showAll = _cbTargetFilter.SelectedIndex == 1;
                    _dgTargets.Items.Filter = showAll ? (Predicate<object>)null :
                        obj => (obj as TargetDoseRow)?.TargetId?.IndexOf("PTV", StringComparison.OrdinalIgnoreCase) >= 0;
                }

                private bool CommitSelections()
                {
                    _dgTargets.CommitEdit(DataGridEditingUnit.Cell, true);
                    _dgTargets.CommitEdit(DataGridEditingUnit.Row, true);
                    _dgOrgans.CommitEdit(DataGridEditingUnit.Cell, true);
                    _dgOrgans.CommitEdit(DataGridEditingUnit.Row, true);

                    if (IsBreast && !_vm.IsLeftSided.HasValue)
                    {
                        MessageBox.Show(_owner,
                            "Please select 'Left Breast' or 'Right Breast' before proceeding.\n\n" +
                            "Laterality is required to orient the asymmetric Virtual Bolus generation.",
                            "Laterality Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }

                    // ----------------------------------------------------------------
                    // v3.0.0.30: Validate BOLUS structure exists if Physical Bolus ticked
                    // ----------------------------------------------------------------
                    if (IsBreast && _vm.HasPhysicalBolus)
                    {
                        var bolusStructure = _ss.Structures.FirstOrDefault(s =>
                            s != null && !s.IsEmpty &&
                            string.Equals(s.DicomType, "BOLUS", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(s.Id, "Bolus_physical", StringComparison.OrdinalIgnoreCase));

                        if (bolusStructure == null)
                        {
                            MessageBox.Show(_owner,
                                "\"Physical Bolus present\" is ticked, but no structure with DICOM type BOLUS was found in this structure set.\n\n" +
                                "Please add the bolus via Insert → New Bolus... in Eclipse before running this script, or untick the Physical Bolus option.",
                                "Missing BOLUS Structure", MessageBoxButton.OK, MessageBoxImage.Error);
                            return false;
                        }
                    }

                    if (_peakProjectedStructures > MAX_STRUCTURES)
                    {
                        MessageBox.Show(
                            $"This operation will exceed the Varian 255 structure limit.\n\n" +
                            $"Existing: {_ss.Structures.Count()}\n" +
                            $"Peak projected: {_peakProjectedStructures}\n" +
                            $"Limit: {MAX_STRUCTURES}\n\n" +
                            "Please untick some options or delete unused structures.",
                            "Structure Limit Exceeded", MessageBoxButton.OK, MessageBoxImage.Error);
                        return false;
                    }

                    var selectedTargets = _vm.TargetDoseRows.Where(r => r.IsSelected).ToList();
                    if (selectedTargets.Count == 0)
                    {
                        MessageBox.Show(_owner, "Please tick at least one target.",
                            "No targets selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }

                    var badDose = selectedTargets
                        .Where(r => !r.ParsedDoseGy.HasValue || r.ParsedDoseGy.Value <= 0)
                        .Select(r => r.TargetId).ToList();
                    if (badDose.Count > 0)
                    {
                        MessageBox.Show(_owner,
                            "Missing or invalid Dose (Gy) for:\n" + string.Join("\n", badDose),
                            "Missing dose", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }

                    var badSuffix = selectedTargets
                        .Where(r => !string.IsNullOrWhiteSpace(r.Suffix) && !IsSafeIdFragment(r.Suffix))
                        .Select(r => $"{r.TargetId}: '{r.Suffix}'").ToList();
                    if (badSuffix.Count > 0)
                    {
                        MessageBox.Show(_owner,
                            "Suffix values with disallowed characters (use letters, digits, _, -, . only):\n\n" +
                            string.Join("\n", badSuffix),
                            "Invalid suffix", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }

                    var dupSuffixes = selectedTargets
                        .Where(r => !string.IsNullOrWhiteSpace(r.Suffix))
                        .GroupBy(r => r.Suffix.Trim(), StringComparer.OrdinalIgnoreCase)
                        .Where(g => g.Select(r => r.ParsedDoseGy.Value).Distinct().Count() > 1)
                        .Select(g => $"Suffix '{g.Key}' used for doses: " +
                                     string.Join(", ", g.Select(r => r.ParsedDoseGy.Value).Distinct()) + " Gy")
                        .ToList();
                    if (dupSuffixes.Count > 0)
                    {
                        MessageBox.Show(_owner,
                            "Same Suffix mapped to different dose levels:\n\n" + string.Join("\n", dupSuffixes),
                            "Duplicate suffixes across doses", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }

                    if (IsBreast)
                    {
                        var blankBolus = selectedTargets
                            .Where(r => string.IsNullOrWhiteSpace(r.BolusMm))
                            .Select(r => r.TargetId).ToList();
                        if (blankBolus.Count > 0)
                        {
                            var result = MessageBox.Show(_owner,
                                "Empty Bolus field – virtual bolus will not be created for:\n\n" +
                                string.Join("\n", blankBolus) + "\n\nProceed without bolus?",
                                "Missing Bolus", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                            if (result == MessageBoxResult.No) return false;
                        }

                        var badBolus = selectedTargets
                            .Where(r => !string.IsNullOrWhiteSpace(r.BolusMm) &&
                                        (!r.ParsedBolusMm.HasValue || r.ParsedBolusMm.Value <= 0))
                            .Select(r => r.TargetId).ToList();
                        if (badBolus.Count > 0)
                        {
                            MessageBox.Show(_owner,
                                "Invalid Bolus thickness (must be > 0 or blank):\n" + string.Join("\n", badBolus),
                                "Invalid Bolus thickness", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return false;
                        }
                    }

                    var badPrv = _vm.OrganRows
                        .Where(r => r.CreatePrv && (!r.ParsedPrvMarginMm.HasValue || r.ParsedPrvMarginMm.Value <= 0))
                        .Select(r => r.OarId).ToList();
                    if (badPrv.Count > 0)
                    {
                        MessageBox.Show(_owner,
                            "PRV ticked but margin ≤ 0 or blank:\n" + string.Join("\n", badPrv),
                            "Invalid PRV margin", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return false;
                    }

                    return true;
                }

                private void ReSortAndTickOrgans()
                {
                    if (!_vm.IsLeftSided.HasValue) return;
                    bool isLeft = _vm.IsLeftSided.Value;

                    foreach (var row in _vm.OrganRows)
                    {
                        bool isPrio = IsIpsilateralPriority(row, isLeft);
                        row.CreateOvl = isPrio;
                        row.CreateOpt = isPrio;
                        row.CreatePrv = false;
                    }

                    _dgOrgans.ItemsSource = null;

                    _vm.OrganRows.Sort((a, b) =>
                    {
                        bool aPrio = IsIpsilateralPriority(a, isLeft);
                        bool bPrio = IsIpsilateralPriority(b, isLeft);
                        if (aPrio && !bPrio) return -1;
                        if (!aPrio && bPrio) return 1;
                        return string.Compare(a.OarId, b.OarId, StringComparison.OrdinalIgnoreCase);
                    });

                    _dgOrgans.ItemsSource = _vm.OrganRows;
                    UpdateStructureCount();
                }

                private void UpdateStructureCount()
                {
                    if (_txtStats == null) return;

                    int existingCount = _ss.Structures.Count();
                    int expectedNew = 0;

                    var selectedTargets = _vm.TargetDoseRows
                        .Where(r => r.IsSelected && r.ParsedDoseGy.HasValue)
                        .ToList();

                    var targetDosePairs = selectedTargets.Select(r => new
                    {
                        r.TargetId,
                        DoseGy = r.ParsedDoseGy.Value,
                        Suffix = (r.Suffix ?? "").Trim(),
                        BolusMm = r.ParsedBolusMm,
                        r.CreateAvoidance
                    }).ToList();

                    var groupKeys = targetDosePairs.Select(x => new { x.DoseGy, x.Suffix }).Distinct().ToList();
                    var doseLevels = groupKeys.Select(k => k.DoseGy).Distinct().ToList();

                    if (selectedTargets.Count > 0)
                    {
                        expectedNew += 1;
                        expectedNew += doseLevels.Count;
                        expectedNew += groupKeys.Count;
                        expectedNew += groupKeys.Count;
                        expectedNew += doseLevels.Count;

                        foreach (var k in groupKeys)
                        {
                            if (targetDosePairs.Any(t =>
                                t.DoseGy == k.DoseGy &&
                                string.Equals(t.Suffix, k.Suffix, StringComparison.OrdinalIgnoreCase) &&
                                t.CreateAvoidance))
                                expectedNew++;
                        }

                        if (targetDosePairs.Any(t => t.BolusMm.GetValueOrDefault() > 0))
                        {
                            expectedNew += 4; // z_Virtual_PTV + z_Virtual_Bolus + z_Virtual_PTV_Opt + Body_new

                            if (_vm.HasPhysicalBolus)
                                expectedNew += 2; // Bolus_physical + Bolus_phys_Opt
                        }

                        expectedNew += doseLevels.Count * 2;
                    }

                    int ovlCount = _vm.OrganRows.Count(r => r.CreateOvl);
                    int optCount = _vm.OrganRows.Count(r => r.CreateOpt);
                    int prvCount = _vm.OrganRows.Count(r => r.CreatePrv);

                    expectedNew += ovlCount * Math.Max(doseLevels.Count, 1);
                    expectedNew += optCount;
                    expectedNew += prvCount;

                    int bufferTemps = Math.Max(5, groupKeys.Count);
                    _peakProjectedStructures = existingCount + expectedNew + bufferTemps;
                    int available = Math.Max(0, MAX_STRUCTURES - _peakProjectedStructures);

                    _txtStats.Text = $"Structure Count: {existingCount} existing + {expectedNew} final " +
                                     $"(+ {bufferTemps} peak temps) = Peak {_peakProjectedStructures} " +
                                     $"/ {MAX_STRUCTURES} limit  (Avail: {available})";

                    _txtStats.Foreground = _peakProjectedStructures > MAX_STRUCTURES
                        ? Brushes.Red
                        : (Brush)_owner.FindResource("AccentCyan");
                }

                // ==================================================================
                // RCC ENGINE — "RCC Optimization Cropping Method" cheat-sheet formulas.
                // Reads Rx from TargetDoseRow (IsSelected + DoseGy) and Max Dose /
                // Nested-sparing from OrganRow (MaxDoseGy + IsSmallOrgan/IsLargeOrgan + NestedSparing),
                // i.e. the exact same rows ticked in the shared grids above.
                // ==================================================================

                // One "label above, rate below" column, used to lay the FALLOFF
                // ZONE section out as a single row with one column per zone.
                private TextBox AddRateColumn(UniformGrid parent, string label, double defaultValue)
                {
                    var col = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 0, 12, 0) };
                    col.Children.Add(new TextBlock
                    {
                        Text = label,
                        FontSize = 11,
                        Foreground = (Brush)_owner.FindResource("TextSecondary"),
                        Margin = new Thickness(0, 0, 0, 3),
                        TextWrapping = TextWrapping.Wrap
                    });
                    var tb = new TextBox
                    {
                        Text = defaultValue.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture),
                        Width = 70,
                        HorizontalAlignment = HorizontalAlignment.Left
                    };
                    tb.LostFocus += (s, e) => RefreshRccPlan();
                    col.Children.Add(tb);
                    parent.Children.Add(col);
                    return tb;
                }

                private double ParseRateOrDefault(TextBox tb, double fallback)
                {
                    return double.TryParse(tb?.Text,
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var v) && v > 0
                        ? v : fallback;
                }

                private static double RccPctDiff(double rxHighOrRef, double comparisonDoseGy) =>
                    (rxHighOrRef - comparisonDoseGy) / rxHighOrRef * 100.0;

                private static double RccCropMm(double pctDiff, double falloffRatePctPerMm) =>
                    falloffRatePctPerMm > 0 ? pctDiff / falloffRatePctPerMm : 0.0;

                // Naming for the RCC Eval/Opt/Sum crop pipeline below. There is
                // exactly one PTV_Opt per target for the whole RCC engine - the
                // OAR max-dose crop, SIB shave and nested §7 crop all reshape the
                // SAME z{target}_Opt structure in place rather than each producing
                // their own "_Opti" result, unlike how the cheat sheet names things.
                // Also distinct from the Generic/Breast pipeline's own
                // "z_PTV_eval_{dose}"/"z_PTV_opt_{dose}"/"z_PTV_opt_sum" naming.
                private static string RccEvalId(string targetId) => TruncId($"z{targetId}_Eval");
                private static string RccOptId(string targetId) => TruncId($"z{targetId}_Opt");
                private static string RccRindId(string targetId) => TruncId($"z{targetId}_Rind");
                private const string RCC_OPT_SUM_ID = "zRCC_Opt_Sum";

                // §7 nested OAR-in-PTV shell, one structure per OAR per level
                // (level 1 = nearest the OAR surface), unioned across every ticked
                // target that OAR overlaps - see BuildRccNestedShells.
                private static string RccNestedShellId(string oarId, int level) => TruncId($"z_{oarId}_in_ptv_hr{level}");

                // Every RCC-created/updated structure is forced to high resolution
                // regardless of the source structures' own resolution, for finer
                // geometric fidelity in the final crops - unlike the rest of the
                // codebase's ConvertToHighResolution calls, which only match an
                // operand's resolution when needed for a boolean op to succeed.
                private static void EnsureRccHighRes(Structure st)
                {
                    if (st != null && !st.IsHighResolution) st.ConvertToHighResolution();
                }

                // Recomputes the RCC plan from current ticks/Rx/MaxDose/zone-rate
                // inputs and repopulates the crop-distance matrix and the advanced
                // SIB/ring/nested preview grid, plus the shared bottom-bar stats line.
                private void RefreshRccPlan()
                {
                    if (_dgRccMatrix == null || _dgRccPlan == null) return;

                    _dgTargets?.CommitEdit(DataGridEditingUnit.Cell, true);
                    _dgTargets?.CommitEdit(DataGridEditingUnit.Row, true);
                    _dgOrgans?.CommitEdit(DataGridEditingUnit.Cell, true);
                    _dgOrgans?.CommitEdit(DataGridEditingUnit.Row, true);

                    _rccPlan = ComputeRccPlan();
                    RefreshRccMatrix();

                    var rows = new List<RccPlanRow>();

                    // Steps 1-2: PTV_Eval / PTV_Opt per ticked target, then PTV_Opt_Sum
                    // (the base Ring1/Ring2 below expand outward from).
                    foreach (var lvl in _rccPlan.Ring1Levels)
                        rows.Add(new RccPlanRow
                        {
                            Category = "PTV Eval/Opt (§1-2)",
                            Source = $"{lvl.Target.TargetId} ∩ (Body-3mm), then +2mm",
                            Zone = "-",
                            PctDiff = 0,
                            CropMm = 0,
                            ResultId = RccOptId(lvl.Target.TargetId)
                        });
                    if (_rccPlan.Ring1Levels.Count > 0)
                        rows.Add(new RccPlanRow
                        {
                            Category = "PTV Opt Sum (§2)",
                            Source = $"union of {_rccPlan.Ring1Levels.Count} PTV_Opt structure(s)",
                            Zone = "-",
                            PctDiff = 0,
                            CropMm = 0,
                            ResultId = RCC_OPT_SUM_ID
                        });

                    // Ring1/Ring2 (and SIB shave below) only run once there are 2+
                    // ticked targets to compare against each other - Ring1Levels has
                    // one entry per ticked target regardless, so gate on count here
                    // to keep the preview honest about what Generate Structure will
                    // actually build for a single-target selection.
                    if (_rccPlan.Ring1Levels.Count >= 2)
                    {
                        foreach (var lvl in _rccPlan.Ring1Levels)
                            rows.Add(new RccPlanRow
                            {
                                Category = "z_Ring_1 (§4)",
                                Source = $"from {lvl.Target.TargetId} (target {_rccPlan.Ring1TargetDoseGy:0.##} Gy = 85% x {_rccPlan.LowestSibRxGy:0.##} Gy)",
                                Zone = "B",
                                PctDiff = 0,
                                CropMm = lvl.CropMm,
                                ResultId = "z_Ring_1"
                            });

                        foreach (var lvl in _rccPlan.Ring2Levels)
                            rows.Add(new RccPlanRow
                            {
                                Category = "z_Ring_2 (§5)",
                                Source = $"from {lvl.Target.TargetId} (target {_rccPlan.Ring2TargetDoseGy:0.##} Gy = 65% x {_rccPlan.LowestSibRxGy:0.##} Gy)",
                                Zone = "C",
                                PctDiff = 0,
                                CropMm = lvl.CropMm,
                                ResultId = "z_Ring_2"
                            });
                    }

                    // Step 5: SIB shave, now on the PTV_Opt structures from §1-2
                    // (not the raw PTV).
                    foreach (var sh in _rccPlan.SibShaves)
                        rows.Add(new RccPlanRow
                        {
                            Category = "SIB shave PTV_Opt (§3)",
                            Source = $"{RccOptId(sh.Low.TargetId)} shaved from {RccOptId(sh.High.TargetId)}",
                            Zone = "B",
                            PctDiff = sh.PctDiff,
                            CropMm = sh.CropMm,
                            ResultId = sh.ResultId
                        });

                    // Step 6: Rind = each target's final PTV_Opt contracted inward
                    // (post-shave boundary for targets shaved in Step 5 above).
                    foreach (var lvl in _rccPlan.Ring1Levels)
                        rows.Add(new RccPlanRow
                        {
                            Category = "Rind (§6)",
                            Source = $"{RccOptId(lvl.Target.TargetId)} contracted {RCC_RIND_MARGIN_MM:0.#}mm",
                            Zone = "-",
                            PctDiff = 0,
                            CropMm = RCC_RIND_MARGIN_MM,
                            ResultId = RccRindId(lvl.Target.TargetId)
                        });

                    foreach (var nr in _rccPlan.NestedRings)
                        rows.Add(new RccPlanRow { Category = "PTV_Opt nested crop (§7 full)", Source = $"{nr.Target.TargetId} minus {nr.Oar.OarId}", Zone = "-", PctDiff = 0, CropMm = 0, ResultId = RccOptId(nr.Target.TargetId) });

                    // Shell count/geometry can only be known once real segment
                    // volumes are booleaned at generation time, so the preview just
                    // names the pattern and step size per OAR rather than guessing N.
                    foreach (var oarGroup in _rccPlan.NestedRings.GroupBy(nr => nr.Oar.OarId, StringComparer.OrdinalIgnoreCase))
                    {
                        var oar = oarGroup.First().Oar;
                        double thickness = oar.ParsedNestedThicknessMm ?? RCC_NESTED_RING_STEP_MM;
                        if (thickness <= 0) thickness = RCC_NESTED_RING_STEP_MM;
                        var targetIds = string.Join(", ", oarGroup.Select(nr => nr.Target.TargetId));
                        rows.Add(new RccPlanRow
                        {
                            Category = "OAR-in-PTV shells (§7)",
                            Source = $"{oar.OarId} ∩ [{targetIds}], stepped {thickness:0.#}mm shells until clear",
                            Zone = "-",
                            PctDiff = 0,
                            CropMm = thickness,
                            ResultId = RccNestedShellId(oar.OarId, 1) + " .. hrN"
                        });
                    }

                    _dgRccPlan.ItemsSource = rows;

                    if (_txtStats != null)
                    {
                        int targetCount = _rccPlan.Ring1Levels.Count;
                        int applicablePairs = _rccPlan.MaxDoseCrops.Count;
                        int targetsToCrop = _rccPlan.MaxDoseCrops.Select(c => c.Target.TargetId).Distinct(StringComparer.OrdinalIgnoreCase).Count();
                        _txtStats.Text = $"{targetCount} target(s): Eval/Opt/Rind always; " +
                                         $"{applicablePairs} OAR max-dose crop(s) on {targetsToCrop} PTV(s); " +
                                         (targetCount >= 2
                                             ? $"{_rccPlan.SibShaves.Count} SIB shave(s) + rings (2+ targets); "
                                             : "no SIB shave/rings (need 2+ targets); ") +
                                         $"{_rccPlan.NestedRings.Count} nested-sparing pair(s).";
                        _txtStats.Foreground = (Brush)_owner.FindResource("AccentCyan");
                    }
                }

                // Rebuilds the read-only "Crop Distance Matrix" - one row per ticked
                // OAR, one column per ticked target, cell = auto-computed crop mm (or
                // "-" if that OAR's Max Dose does not require sparing this target's Rx).
                private void RefreshRccMatrix()
                {
                    double zoneASmall = ParseRateOrDefault(_txtRccZoneASmall, RCC_ZONE_A_SMALL_DEFAULT_PCT_PER_MM);
                    double zoneALarge = ParseRateOrDefault(_txtRccZoneALarge, RCC_ZONE_A_LARGE_DEFAULT_PCT_PER_MM);
                    var tickedTargets = _vm.TargetDoseRows
                        .Where(r => r.IsSelected && r.ParsedDoseGy.HasValue && r.ParsedDoseGy.Value > 0)
                        .ToList();
                    // An OAR is included the moment it has a valid Max Dose - no
                    // separate tick anymore.
                    var tickedOars = _vm.OrganRows.Where(r => r.ParsedMaxDoseGy.HasValue).ToList();

                    var matrixRows = new List<RccMatrixRow>();
                    foreach (var oar in tickedOars)
                    {
                        double zoneA = oar.IsLargeOrgan ? zoneALarge : zoneASmall;
                        var row = new RccMatrixRow
                        {
                            OarId = oar.OarId,
                            MaxDoseDisplay = oar.ParsedMaxDoseGy.HasValue
                                ? oar.ParsedMaxDoseGy.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) + " Gy"
                                : "(set)",
                            CropDisplay = new string[tickedTargets.Count]
                        };

                        for (int i = 0; i < tickedTargets.Count; i++)
                        {
                            var t = tickedTargets[i];
                            if (!oar.ParsedMaxDoseGy.HasValue) { row.CropDisplay[i] = "–"; continue; }

                            double rx = t.ParsedDoseGy.Value;
                            double oarMax = oar.ParsedMaxDoseGy.Value;
                            if (oarMax >= rx) { row.CropDisplay[i] = "–"; continue; }

                            double pctDiff = RccPctDiff(rx, oarMax);
                            double cropMm = RccCropMm(pctDiff, zoneA);
                            row.CropDisplay[i] = cropMm > 0
                                ? cropMm.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + " mm"
                                : "–";
                        }
                        matrixRows.Add(row);
                    }

                    _dgRccMatrix.Columns.Clear();
                    _dgRccMatrix.Columns.Add(new DataGridTextColumn
                    {
                        Header = "OAR",
                        Binding = new Binding("OarId"),
                        IsReadOnly = true,
                        Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                        MinWidth = 90,
                        ElementStyle = (Style)_owner.FindResource(typeof(TextBlock))
                    });
                    _dgRccMatrix.Columns.Add(new DataGridTextColumn
                    {
                        Header = "Max Dose",
                        Binding = new Binding("MaxDoseDisplay"),
                        IsReadOnly = true,
                        Width = 80,
                        ElementStyle = (Style)_owner.FindResource(typeof(TextBlock))
                    });
                    for (int i = 0; i < tickedTargets.Count; i++)
                    {
                        string tid = tickedTargets[i].TargetId;
                        string headerName = tid.Length > 10 ? tid.Substring(0, 10) + ".." : tid;
                        _dgRccMatrix.Columns.Add(new DataGridTextColumn
                        {
                            Header = headerName + " crop",
                            Binding = new Binding($"CropDisplay[{i}]"),
                            IsReadOnly = true,
                            Width = 90,
                            ElementStyle = (Style)_owner.FindResource(typeof(TextBlock))
                        });
                    }
                    _dgRccMatrix.ItemsSource = matrixRows;
                }

                // Pure formula engine - no ESAPI calls here, only the PDF cheat-sheet math.
                private RccPlan ComputeRccPlan()
                {
                    var plan = new RccPlan();
                    double zoneASmall = ParseRateOrDefault(_txtRccZoneASmall, RCC_ZONE_A_SMALL_DEFAULT_PCT_PER_MM);
                    double zoneALarge = ParseRateOrDefault(_txtRccZoneALarge, RCC_ZONE_A_LARGE_DEFAULT_PCT_PER_MM);
                    double zoneB = ParseRateOrDefault(_txtRccZoneB, RCC_ZONE_B_DEFAULT_PCT_PER_MM);
                    double zoneC = ParseRateOrDefault(_txtRccZoneC, RCC_ZONE_C_DEFAULT_PCT_PER_MM);

                    var selTargets = _vm.TargetDoseRows
                        .Where(r => r.IsSelected && r.ParsedDoseGy.HasValue && r.ParsedDoseGy.Value > 0)
                        .ToList();

                    // Section 2: zPTV Opti (strict Max-Dose OAR sparing crop), Zone A
                    // (small-organ or large-organ rate, per that OAR's own tick).
                    foreach (var oar in _vm.OrganRows.Where(o => o.ParsedMaxDoseGy.HasValue))
                    {
                        double zoneA = oar.IsLargeOrgan ? zoneALarge : zoneASmall;
                        foreach (var t in selTargets)
                        {
                            double rx = t.ParsedDoseGy.Value;
                            double oarMax = oar.ParsedMaxDoseGy.Value;
                            if (oarMax >= rx) continue;

                            double pctDiff = RccPctDiff(rx, oarMax);
                            double cropMm = RccCropMm(pctDiff, zoneA);
                            if (cropMm <= 0) continue;

                            plan.MaxDoseCrops.Add(new RccMaxDoseCrop
                            {
                                Target = t,
                                Oar = oar,
                                PctDiff = pctDiff,
                                CropMm = cropMm,
                                ResultId = RccOptId(t.TargetId)
                            });
                        }
                    }

                    // Sections 3-5: SIB ladder shave + variable rings, Zone B / Zone C.
                    var sibSorted = selTargets.OrderByDescending(t => t.ParsedDoseGy.Value).ToList();
                    for (int i = 0; i < sibSorted.Count - 1; i++)
                    {
                        var high = sibSorted[i];
                        var low = sibSorted[i + 1];
                        double rxHigh = high.ParsedDoseGy.Value;
                        double rxLow = low.ParsedDoseGy.Value;
                        double pctDiff = RccPctDiff(rxHigh, rxLow);
                        double cropMm = RccCropMm(pctDiff, zoneB);

                        plan.SibShaves.Add(new RccSibShave
                        {
                            High = high,
                            Low = low,
                            PctDiff = pctDiff,
                            CropMm = cropMm,
                            ResultId = RccOptId(low.TargetId)
                        });
                    }

                    if (sibSorted.Count > 0)
                    {
                        double lowestRx = sibSorted[sibSorted.Count - 1].ParsedDoseGy.Value;
                        plan.LowestSibRxGy = lowestRx;
                        plan.Ring1TargetDoseGy = RCC_RING1_ISO_FRACTION * lowestRx;
                        plan.Ring2TargetDoseGy = RCC_RING2_ISO_FRACTION * lowestRx;

                        foreach (var t in sibSorted)
                        {
                            double rx = t.ParsedDoseGy.Value;

                            double pctDiff1 = (rx - plan.Ring1TargetDoseGy) / rx * 100.0;
                            double crop1 = Math.Max(RCC_RING1_MIN_CROP_MM, RccCropMm(pctDiff1, zoneB));
                            plan.Ring1Levels.Add(new RccRingLevel { Target = t, CropMm = crop1 });

                            double pctDiff2 = (rx - plan.Ring2TargetDoseGy) / rx * 100.0;
                            double crop2 = Math.Max(0.0, RccCropMm(pctDiff2, zoneC));
                            plan.Ring2Levels.Add(new RccRingLevel { Target = t, CropMm = crop2 });
                        }
                    }

                    // Section 7: iterative nested-ring sparing (OAR mean-dose overlap).
                    foreach (var oar in _vm.OrganRows.Where(o => o.NestedSparing))
                        foreach (var t in selTargets)
                            plan.NestedRings.Add(new RccNestedRing { Target = t, Oar = oar });

                    return plan;
                }

                // ------------------------------------------------------------
                // Pre-flight check for Generate Structure: returns a human-
                // readable list of missing/invalid required input (empty when
                // everything needed - External/Body, >=1 ticked target with a
                // valid Rx - is present). OAR Max Dose is optional (Eval/Opt/
                // Rind, and Ring/SIB for 2+ targets, don't need any OAR at all)
                // but any Max Dose that IS entered must be a valid number.
                // ------------------------------------------------------------
                private List<string> ValidateRccGenerateInputs(Structure ext)
                {
                    var missing = new List<string>();

                    if (ext == null || ext.IsEmpty)
                        missing.Add("External/Body structure");

                    var tickedTargets = _vm.TargetDoseRows.Where(r => r.IsSelected).ToList();
                    if (tickedTargets.Count == 0)
                    {
                        missing.Add("At least one ticked target");
                    }
                    else
                    {
                        var badRx = tickedTargets
                            .Where(r => !r.ParsedDoseGy.HasValue || r.ParsedDoseGy.Value <= 0)
                            .Select(r => r.TargetId).ToList();
                        if (badRx.Count > 0)
                            missing.Add("Rx (Gy) for: " + string.Join(", ", badRx));
                    }

                    // OAR Max Dose is optional; an entered value just has to be valid.
                    var badMax = _vm.OrganRows
                        .Where(r => !string.IsNullOrWhiteSpace(r.MaxDoseGy) &&
                                    (!r.ParsedMaxDoseGy.HasValue || r.ParsedMaxDoseGy.Value <= 0))
                        .Select(r => r.OarId).ToList();
                    if (badMax.Count > 0)
                        missing.Add("Max Dose (Gy) for: " + string.Join(", ", badMax));

                    return missing;
                }

                // Single vs multi-target (SIB) are one process now, both starting
                // from Eval/Opt. Missing/invalid input blocks the run (alert, stays
                // open so it can be fixed). Once generation actually runs, the
                // dialog closes right away regardless of outcome - a creation error
                // still shows a summary of what succeeded/failed first, then closes.
                private void DoRccGenerateStructure()
                {
                    RefreshRccPlan();

                    var ext = _vm.SelectedExternal ?? FindExternalFallback(_ss);

                    var missing = ValidateRccGenerateInputs(ext);
                    if (missing.Count > 0)
                    {
                        MessageBox.Show(_owner,
                            "Cannot generate - missing/invalid input:\n\n" +
                            string.Join("\n", missing.Select(m => "- " + m)),
                            "Missing Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    var plan = _rccPlan;
                    // Validation above already guarantees at least one ticked
                    // target with a valid Rx, so pipelineTargets is never empty here.
                    var pipelineTargets = _vm.TargetDoseRows
                        .Where(r => r.IsSelected && r.ParsedDoseGy.HasValue && r.ParsedDoseGy.Value > 0)
                        .ToList();

                    var fb = new SliceRecontourFallback();
                    var created = new List<string>();
                    var errors = new List<string>();

                    // Steps 1-2: PTV_Eval / PTV_Opt per target.
                    var optByTarget = BuildRccEvalOptStructures(pipelineTargets, ext, created, errors);

                    // §2: OAR max-dose crop, applied to Opt in place - required for
                    // both the single-target and multi-target pipelines alike.
                    ApplyRccMaxDoseCropToOpt(plan, optByTarget, ext, created, errors);

                    // PTV_Opt_Sum, built from the (now OAR-cropped) Opt structures.
                    var optSum = BuildRccOptSum(optByTarget, ext, created, errors);

                    if (pipelineTargets.Count >= 2)
                    {
                        // Steps 3-4: Ring1 (Zone B) then Ring2 (Zone C, outside
                        // ring1) - only meaningful once there's another PTV to
                        // fall off around.
                        var ring1St = BuildRccRing1(plan, optSum, optByTarget, ext, created, errors);
                        BuildRccRing2(plan, optSum, optByTarget, ext, ring1St, created, errors);

                        // Step 5: SIB-shave each lower-dose target's Opt from the
                        // next higher-dose target's Opt, in place.
                        ApplyRccSibShaveToOpt(plan, optByTarget, ext, created, errors);
                    }

                    // ---- Section 7: nested OAR-in-PTV sparing rings. The "full
                    // crop from the overlapping OAR" is now just another
                    // mutation of the SAME z{target}_Opt used throughout RCC -
                    // there is only ever one PTV_Opt per target, never a second
                    // "_Opti" structure. ----
                    if (plan.NestedRings.Count > 0)
                    {
                        // Step A: crop each (target, OAR) pair's PTV_Opt in place.
                        foreach (var nr in plan.NestedRings)
                        {
                            try
                            {
                                using (var tg = new TempGuard(_ss))
                                {
                                    var oarSt = _ss.Structures.FirstOrDefault(s => !s.IsEmpty && string.Equals(s.Id, nr.Oar.OarId, StringComparison.OrdinalIgnoreCase));
                                    if (oarSt == null || !optByTarget.TryGetValue(nr.Target.TargetId, out var targetOptSt))
                                    {
                                        errors.Add($"{nr.Target.TargetId}/{nr.Oar.OarId}: source structure missing");
                                        continue;
                                    }

                                    var optCroppedSeg = SafeBoolean(_ss, targetOptSt.SegmentVolume, oarSt.SegmentVolume, BoolOp.Sub,
                                        targetOptSt, oarSt, targetOptSt.Id, fb, $"RccNestedOptCrop_{targetOptSt.Id}", tg);
                                    optCroppedSeg = SafeBoolean(_ss, optCroppedSeg, ext.SegmentVolume, BoolOp.And,
                                        null, ext, targetOptSt.Id, fb, $"RccNestedOptCrop_{targetOptSt.Id}_CapExt", tg);

                                    if (AssignSegmentSafely(targetOptSt, optCroppedSeg))
                                        created.Add($"{targetOptSt.Id}  <-  nested OAR crop: {nr.Oar.OarId}");
                                    else
                                        errors.Add($"{targetOptSt.Id}: empty result after nested OAR crop");
                                }
                            }
                            catch (Exception ex) { errors.Add($"{nr.Target.TargetId}/{nr.Oar.OarId}: {ex.Message}"); }
                        }

                        // Step B: z_oar_in_ptv_hr# shells, one level set per OAR
                        // (unioned across every ticked target that OAR overlaps),
                        // stepped outward from the raw OAR/target geometry - not the
                        // now-further-cropped Opt boundary from Step A above.
                        foreach (var oarGroup in plan.NestedRings.GroupBy(nr => nr.Oar.OarId, StringComparer.OrdinalIgnoreCase))
                        {
                            var oar = oarGroup.First().Oar;
                            var targetIds = oarGroup.Select(nr => nr.Target.TargetId)
                                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                            BuildRccNestedShells(oar, targetIds, fb, created, errors);
                        }
                    }

                    // No smoothing here - PTV_Opt is exactly PTV_Eval cropped by the
                    // formulas above (§2 OAR max-dose, SIB shave, §7 nested), never
                    // expanded/smoothed past that boundary. SMOOTH_OPT_TARGET/
                    // SMOOTH_MM remain in use by the separate Generic/Breast Opto
                    // pipeline only.

                    // Step 6: Rind = each target's FINAL PTV_Opt contracted inward.
                    BuildRccRind(optByTarget, ext, created, errors);

                    if (errors.Count > 0)
                    {
                        var summary = new StringBuilder();
                        summary.AppendLine($"Generated/updated {created.Count} structure(s):");
                        foreach (var id in created) summary.AppendLine($"  {id}");
                        summary.AppendLine();
                        summary.AppendLine($"Errors ({errors.Count}):");
                        foreach (var err in errors) summary.AppendLine($"  {err}");
                        MessageBox.Show(_owner, summary.ToString(), "RCC Generate Structure", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }

                    // Close immediately after Generate Structure runs, whether it
                    // finished clean or with errors (the error summary above is
                    // shown first so nothing silently fails).
                    _owner.CancelDialog();
                }

                // ==============================================================
                // RCC UNIFIED GENERATE-STRUCTURE PIPELINE HELPERS
                // Single vs multi-target (SIB) are the same process, both
                // starting from Eval/Opt: PTV_Eval -> PTV_Opt -> OAR max-dose
                // crop (applied to Opt, required for both cases alike) ->
                // PTV_Opt_Sum -> [2+ targets only: Ring1 -> Ring2 -> SIB shave]
                // -> Rind. One ticked target still gets Eval/Opt/OAR-crop/Rind;
                // it just has no "other" PTV to build a ring or SIB-shave
                // against, so those two steps are skipped for it (see
                // DoRccGenerateStructure above, which calls these in order).
                // Nested OAR-in-PTV sparing (§7) is unaffected by this pipeline
                // (it works off the raw ticked target, per the cheat sheet's
                // own §7 wording) and lives inside DoRccGenerateStructure itself.
                // ==============================================================

                // Step 1 + Step 2: for every ticked target, PTV_Eval = target
                // cropped out of the body by 3mm, then PTV_Opt = PTV_Eval, capped
                // to Body (no expansion - Opt starts out identical to Eval and is
                // reshaped from there by every later crop). Returns the per-target
                // Opt structures; call BuildRccOptSum separately (after any further
                // Opt edits, e.g. the OAR max-dose crop) to (re)build PTV_Opt_Sum.
                private Dictionary<string, Structure> BuildRccEvalOptStructures(
                    List<TargetDoseRow> targets, Structure ext,
                    List<string> created, List<string> errors)
                {
                    var fb = new SliceRecontourFallback();
                    var optByTarget = new Dictionary<string, Structure>(StringComparer.OrdinalIgnoreCase);
                    if (targets.Count == 0) return optByTarget;

                    SegmentVolume bodyMinus3;
                    using (var tg0 = new TempGuard(_ss))
                    {
                        var extMinus3 = SafeMargin(ext.SegmentVolume, -BODY_CONTRACT_MM);
                        bodyMinus3 = SafeBoolean(_ss, extMinus3, ext.SegmentVolume, BoolOp.And,
                            ext, ext, null, fb, "RccBodyMinus3", tg0);
                    }

                    foreach (var row in targets)
                    {
                        var targetSt = _ss.Structures.FirstOrDefault(s =>
                            !s.IsEmpty && string.Equals(s.Id, row.TargetId, StringComparison.OrdinalIgnoreCase));
                        if (targetSt == null) { errors.Add($"{row.TargetId}: target structure missing"); continue; }

                        try
                        {
                            using (var tg = new TempGuard(_ss))
                            {
                                // Step 1: PTV_Eval = target ∩ (Body - 3mm) ∩ Body
                                string evalId = RccEvalId(row.TargetId);
                                var evalSeg = SafeBoolean(_ss, targetSt.SegmentVolume, bodyMinus3, BoolOp.And,
                                    targetSt, ext, evalId, fb, $"RccEval_{row.TargetId}_AndBodyMinus3", tg);
                                evalSeg = SafeBoolean(_ss, evalSeg, ext.SegmentVolume, BoolOp.And,
                                    null, ext, evalId, fb, $"RccEval_{row.TargetId}_AndExt", tg);

                                var evalSt = GetOrCreate(_ss, "PTV", evalId);
                                EnsureRccHighRes(evalSt);
                                if (!AssignSegmentSafely(evalSt, evalSeg))
                                {
                                    _ss.RemoveStructure(evalSt);
                                    errors.Add($"{evalId}: empty result");
                                    continue;
                                }
                                evalSt.Color = Colors.Blue;
                                created.Add(evalId);

                                // Step 2: PTV_Opt = PTV_Eval, capped to Body (no expansion -
                                // every later crop, e.g. OAR max-dose/SIB/§7, reshapes this
                                // same structure starting from the Eval boundary itself).
                                string optId = RccOptId(row.TargetId);
                                var optSeg = SafeBoolean(_ss, evalSt.SegmentVolume, ext.SegmentVolume, BoolOp.And,
                                    evalSt, ext, optId, fb, $"RccOpt_{row.TargetId}_CapExt", tg);

                                var optSt = GetOrCreate(_ss, "PTV", optId);
                                EnsureRccHighRes(optSt);
                                if (!AssignSegmentSafely(optSt, optSeg))
                                {
                                    _ss.RemoveStructure(optSt);
                                    errors.Add($"{optId}: empty result");
                                    continue;
                                }
                                optSt.Color = Colors.Red;
                                created.Add(optId);

                                optByTarget[row.TargetId] = optSt;
                            }
                        }
                        catch (Exception ex) { errors.Add($"{row.TargetId}: {ex.Message}"); }
                    }

                    return optByTarget;
                }

                // PTV_Opt_Sum = union of the current per-target PTV_Opt structures,
                // capped to Body. Called after Eval/Opt (and again after the OAR
                // max-dose crop updates those same Opt structures in place) so the
                // sum - and the Ring1/Ring2 base derived from it - always reflects
                // the latest Opt boundaries.
                private Structure BuildRccOptSum(
                    Dictionary<string, Structure> optByTarget, Structure ext,
                    List<string> created, List<string> errors)
                {
                    if (optByTarget.Count == 0) return null;

                    var fb = new SliceRecontourFallback();
                    using (var tg = new TempGuard(_ss))
                    {
                        var optTemps = optByTarget.Values
                            .Select(s => tg.Add(CreateTempFromSegment(_ss, s.SegmentVolume, "zRCC_OptTmp")))
                            .ToList();
                        var unionSt = tg.Add(UnionManyToTemp(_ss, optTemps, fb, "zRCC_OptSumU", "RccOptSumUnion", tg));
                        if (unionSt == null) return null;

                        var sumSeg = SafeBoolean(_ss, unionSt.SegmentVolume, ext.SegmentVolume, BoolOp.And,
                            unionSt, ext, RCC_OPT_SUM_ID, fb, "RccOptSum_CapExt", tg);

                        var sumSt = GetOrCreate(_ss, "PTV", RCC_OPT_SUM_ID);
                        EnsureRccHighRes(sumSt);
                        if (AssignSegmentSafely(sumSt, sumSeg))
                        {
                            sumSt.Color = Colors.Red;
                            created.Add(RCC_OPT_SUM_ID);
                            return sumSt;
                        }

                        _ss.RemoveStructure(sumSt);
                        errors.Add($"{RCC_OPT_SUM_ID}: empty result");
                        return null;
                    }
                }

                // §2: OAR max-dose crop, applied directly to each target's PTV_Opt
                // (in place) rather than the raw PTV - required for both the
                // single-target and multi-target (SIB) pipelines alike. Reuses
                // plan.MaxDoseCrops (same numbers the matrix shows) grouped by
                // target; each qualifying OAR is expanded by its own Zone A
                // (small/large) crop distance, unioned, and subtracted from that
                // target's Opt.
                private void ApplyRccMaxDoseCropToOpt(
                    RccPlan plan, Dictionary<string, Structure> optByTarget, Structure ext,
                    List<string> created, List<string> errors)
                {
                    var fb = new SliceRecontourFallback();
                    foreach (var grp in plan.MaxDoseCrops.GroupBy(c => c.Target.TargetId, StringComparer.OrdinalIgnoreCase))
                    {
                        if (!optByTarget.TryGetValue(grp.Key, out var optSt))
                        {
                            errors.Add($"{grp.Key}: PTV_Opt missing, OAR max-dose crop not applied");
                            continue;
                        }

                        try
                        {
                            using (var tg = new TempGuard(_ss))
                            {
                                var expandedOarTemps = new List<Structure>();
                                var pairLabels = new List<string>();

                                foreach (var c in grp)
                                {
                                    var oarSt = _ss.Structures.FirstOrDefault(s => !s.IsEmpty && string.Equals(s.Id, c.Oar.OarId, StringComparison.OrdinalIgnoreCase));
                                    if (oarSt == null) { errors.Add($"{optSt.Id}: OAR '{c.Oar.OarId}' missing"); continue; }

                                    var expanded = SafeMargin(oarSt.SegmentVolume, c.CropMm);
                                    expandedOarTemps.Add(tg.Add(CreateTempFromSegment(_ss, expanded, "zRCC_MDExpOar")));
                                    pairLabels.Add($"{c.Oar.OarId}({c.CropMm:0.0}mm)");
                                }
                                if (expandedOarTemps.Count == 0) continue;

                                var oarsUnionSt = tg.Add(UnionManyToTemp(_ss, expandedOarTemps, fb, "zRCC_MDOarsU", $"RccMaxDoseOarsUnion_{grp.Key}", tg));
                                if (oarsUnionSt == null)
                                {
                                    errors.Add($"{optSt.Id}: failed to build expanded-OAR union for {string.Join(", ", pairLabels)} - crop not applied");
                                    continue;
                                }

                                var croppedSeg = SafeBoolean(_ss, optSt.SegmentVolume, oarsUnionSt.SegmentVolume, BoolOp.Sub,
                                    optSt, oarsUnionSt, optSt.Id, fb, $"RccMaxDoseCrop_{grp.Key}_Sub", tg);
                                croppedSeg = SafeBoolean(_ss, croppedSeg, ext.SegmentVolume, BoolOp.And,
                                    null, ext, optSt.Id, fb, $"RccMaxDoseCrop_{grp.Key}_CapExt", tg);

                                if (AssignSegmentSafely(optSt, croppedSeg))
                                    created.Add($"{optSt.Id}  <-  OAR max-dose crop: {string.Join(", ", pairLabels)}");
                                else
                                    errors.Add($"{optSt.Id}: empty result after OAR max-dose crop");
                            }
                        }
                        catch (Exception ex) { errors.Add($"{optSt.Id} OAR max-dose crop: {ex.Message}"); }
                    }
                }

                // Step 3: Ring1, 1cm (10mm) thick, built from PTV_Opt_Sum. The gap
                // (offset from PTV_Opt_Sum to ring1's inner edge) uses the LOWEST
                // ticked Rx's own Ring1 crop distance (Zone B, 85% isodose formula -
                // the smallest/most conservative gap, since the lowest-dose PTV needs
                // the least separation). Every OTHER ticked target's own (larger) Ring1
                // crop distance is then applied to ITS OWN PTV_Opt and subtracted back
                // out, so higher-dose targets get the extra clearance their own
                // formula calls for instead of ring1 encroaching on them.
                private Structure BuildRccRing1(
                    RccPlan plan, Structure optSum, Dictionary<string, Structure> optByTarget,
                    Structure ext, List<string> created, List<string> errors)
                {
                    if (optSum == null || plan.Ring1Levels.Count == 0) return null;

                    // Ring1Levels is built (in ComputeRccPlan) by iterating targets in
                    // descending-Rx order, so the last entry is the lowest-Rx target.
                    var lowestLevel = plan.Ring1Levels[plan.Ring1Levels.Count - 1];
                    var fb = new SliceRecontourFallback();

                    try
                    {
                        using (var tg = new TempGuard(_ss))
                        {
                            var baseSeg = SafeMargin(optSum.SegmentVolume, +lowestLevel.CropMm);
                            var outerSeg = SafeMargin(baseSeg, +RING_OUTER_EXPAND_MM);
                            var baseSt = tg.Add(CreateTempFromSegment(_ss, baseSeg, "zRCC_R1Base"));

                            var ringSeg = SafeBoolean(_ss, outerSeg, baseSeg, BoolOp.Sub,
                                null, baseSt, "z_Ring_1", fb, "Ring1_OuterMinusBase", tg);

                            for (int i = 0; i < plan.Ring1Levels.Count - 1; i++)
                            {
                                var lvl = plan.Ring1Levels[i];
                                if (!optByTarget.TryGetValue(lvl.Target.TargetId, out var higherOpt)) continue;
                                var higherExpanded = SafeMargin(higherOpt.SegmentVolume, lvl.CropMm);
                                var higherSt = tg.Add(CreateTempFromSegment(_ss, higherExpanded, "zRCC_R1HighExp"));
                                ringSeg = SafeBoolean(_ss, ringSeg, higherExpanded, BoolOp.Sub,
                                    null, higherSt, "z_Ring_1", fb, $"Ring1_Sub_{lvl.Target.TargetId}", tg);
                            }

                            ringSeg = SafeBoolean(_ss, ringSeg, ext.SegmentVolume, BoolOp.And,
                                null, ext, "z_Ring_1", fb, "Ring1_CapExt", tg);

                            var st = GetOrCreate(_ss, "CONTROL", "z_Ring_1");
                            EnsureRccHighRes(st);
                            if (AssignSegmentSafely(st, ringSeg))
                            {
                                st.Color = Colors.MediumPurple;
                                created.Add("z_Ring_1");
                                return st;
                            }
                            _ss.RemoveStructure(st);
                            errors.Add("z_Ring_1: empty result");
                            return null;
                        }
                    }
                    catch (Exception ex) { errors.Add($"z_Ring_1: {ex.Message}"); return null; }
                }

                // Step 4: Ring2, 1cm thick, using the same process as Ring1 but with
                // Zone C / 65% isodose gaps (larger than Zone B, so ring2 naturally
                // sits further out). Also explicitly subtracts ring1 so "outside of
                // ring1" holds even for unusual zone-rate inputs.
                private void BuildRccRing2(
                    RccPlan plan, Structure optSum, Dictionary<string, Structure> optByTarget,
                    Structure ext, Structure ring1St, List<string> created, List<string> errors)
                {
                    if (optSum == null || plan.Ring2Levels.Count == 0) return;

                    var lowestLevel = plan.Ring2Levels[plan.Ring2Levels.Count - 1];
                    var fb = new SliceRecontourFallback();

                    try
                    {
                        using (var tg = new TempGuard(_ss))
                        {
                            var baseSeg = SafeMargin(optSum.SegmentVolume, +lowestLevel.CropMm);
                            var outerSeg = SafeMargin(baseSeg, +RING_OUTER_EXPAND_MM);
                            var baseSt = tg.Add(CreateTempFromSegment(_ss, baseSeg, "zRCC_R2Base"));

                            var ringSeg = SafeBoolean(_ss, outerSeg, baseSeg, BoolOp.Sub,
                                null, baseSt, "z_Ring_2", fb, "Ring2_OuterMinusBase", tg);

                            for (int i = 0; i < plan.Ring2Levels.Count - 1; i++)
                            {
                                var lvl = plan.Ring2Levels[i];
                                if (!optByTarget.TryGetValue(lvl.Target.TargetId, out var higherOpt)) continue;
                                var higherExpanded = SafeMargin(higherOpt.SegmentVolume, lvl.CropMm);
                                var higherSt = tg.Add(CreateTempFromSegment(_ss, higherExpanded, "zRCC_R2HighExp"));
                                ringSeg = SafeBoolean(_ss, ringSeg, higherExpanded, BoolOp.Sub,
                                    null, higherSt, "z_Ring_2", fb, $"Ring2_Sub_{lvl.Target.TargetId}", tg);
                            }

                            if (ring1St != null)
                            {
                                var ring1Tmp = tg.Add(CreateTempFromSegment(_ss, ring1St.SegmentVolume, "zRCC_R1Clone"));
                                ringSeg = SafeBoolean(_ss, ringSeg, ring1St.SegmentVolume, BoolOp.Sub,
                                    null, ring1Tmp, "z_Ring_2", fb, "Ring2_SubRing1", tg);
                            }

                            ringSeg = SafeBoolean(_ss, ringSeg, ext.SegmentVolume, BoolOp.And,
                                null, ext, "z_Ring_2", fb, "Ring2_CapExt", tg);

                            var st = GetOrCreate(_ss, "CONTROL", "z_Ring_2");
                            EnsureRccHighRes(st);
                            if (AssignSegmentSafely(st, ringSeg))
                            {
                                st.Color = Colors.SlateBlue;
                                created.Add("z_Ring_2");
                            }
                            else
                            {
                                _ss.RemoveStructure(st);
                                errors.Add("z_Ring_2: empty result");
                            }
                        }
                    }
                    catch (Exception ex) { errors.Add($"z_Ring_2: {ex.Message}"); }
                }

                // Step 5 (2+ targets only): SIB shave, applied IN PLACE to each
                // low-dose target's own PTV_Opt (same z{target}_Opt structure from
                // Steps 1-2/OAR-crop) rather than creating a separate result
                // structure - PTV_Opt is the one evolving structure per target, and
                // downstream steps (Rind, and any Ring1/Ring2 already built) just
                // read optByTarget again to see the latest boundary. plan.SibShaves
                // is ordered highest-to-lowest Rx, so by the time a target is used
                // as the "High" side of one pair it has already been shaved (if
                // applicable) as the "Low" side of the previous pair - the shave
                // cascades correctly through the whole dose ladder.
                private void ApplyRccSibShaveToOpt(
                    RccPlan plan, Dictionary<string, Structure> optByTarget, Structure ext,
                    List<string> created, List<string> errors)
                {
                    var fb = new SliceRecontourFallback();

                    foreach (var sh in plan.SibShaves)
                    {
                        if (!optByTarget.TryGetValue(sh.Low.TargetId, out var lowOptSt) ||
                            !optByTarget.TryGetValue(sh.High.TargetId, out var highOptSt))
                        {
                            errors.Add($"{sh.Low.TargetId}: PTV_Opt source missing for SIB shave");
                            continue;
                        }

                        try
                        {
                            using (var tg = new TempGuard(_ss))
                            {
                                var expandedHigh = SafeMargin(highOptSt.SegmentVolume, sh.CropMm);
                                var expandedHighSt = tg.Add(CreateTempFromSegment(_ss, expandedHigh, "zRCC_ShaveHigh"));

                                var shavedSeg = SafeBoolean(_ss, lowOptSt.SegmentVolume, expandedHigh, BoolOp.Sub,
                                    lowOptSt, expandedHighSt, lowOptSt.Id, fb, $"RccShaveOpt_{lowOptSt.Id}_Sub", tg);
                                shavedSeg = SafeBoolean(_ss, shavedSeg, ext.SegmentVolume, BoolOp.And,
                                    null, ext, lowOptSt.Id, fb, $"RccShaveOpt_{lowOptSt.Id}_CapExt", tg);

                                if (AssignSegmentSafely(lowOptSt, shavedSeg))
                                    created.Add($"{lowOptSt.Id}  <-  SIB shave from {highOptSt.Id} ({sh.CropMm:0.0}mm)");
                                else
                                    errors.Add($"{lowOptSt.Id}: empty result after SIB shave");
                            }
                        }
                        catch (Exception ex) { errors.Add($"{lowOptSt.Id} SIB shave: {ex.Message}"); }
                    }
                }

                // Step 6: Rind = each target's final PTV_Opt contracted inward by
                // RCC_RIND_MARGIN_MM (a plain negative margin, not a boolean shell).
                // optByTarget already reflects the OAR-crop and (if applicable)
                // SIB-shave state since both are applied in place, so this always
                // reads the final boundary regardless of how many targets are ticked.
                private void BuildRccRind(
                    Dictionary<string, Structure> finalOptByTarget, Structure ext,
                    List<string> created, List<string> errors)
                {
                    var fb = new SliceRecontourFallback();
                    foreach (var kvp in finalOptByTarget)
                    {
                        string targetId = kvp.Key;
                        var optSt = kvp.Value;
                        try
                        {
                            using (var tg = new TempGuard(_ss))
                            {
                                string rindId = RccRindId(targetId);
                                var rindSeg = SafeMargin(optSt.SegmentVolume, -RCC_RIND_MARGIN_MM);
                                rindSeg = SafeBoolean(_ss, rindSeg, ext.SegmentVolume, BoolOp.And,
                                    optSt, ext, rindId, fb, $"RccRind_{targetId}_CapExt", tg);

                                var st = GetOrCreate(_ss, "CONTROL", rindId);
                                EnsureRccHighRes(st);
                                if (AssignSegmentSafely(st, rindSeg))
                                {
                                    st.Color = Color.FromRgb(0, 200, 140);
                                    created.Add(rindId);
                                }
                                else
                                {
                                    _ss.RemoveStructure(st);
                                    errors.Add($"{rindId}: empty result");
                                }
                            }
                        }
                        catch (Exception ex) { errors.Add($"z{targetId}_Rind: {ex.Message}"); }
                    }
                }

                // Builds the §7 nested OAR-in-PTV shells for one OAR against every
                // ticked target it overlaps. Shell level k covers the band from
                // (k-1)*thickness to k*thickness mm inward from the OAR surface,
                // clipped to each target and unioned across all of that OAR's
                // targets into one z_oar_in_ptv_hr{k} structure. Stops once a level
                // no longer overlaps any target, so the number of shells created
                // follows the actual OAR/target geometry rather than a fixed count.
                private void BuildRccNestedShells(
                    OrganRow oar, List<string> targetIds,
                    SliceRecontourFallback fb, List<string> created, List<string> errors)
                {
                    var oarSt = _ss.Structures.FirstOrDefault(s => !s.IsEmpty && string.Equals(s.Id, oar.OarId, StringComparison.OrdinalIgnoreCase));
                    if (oarSt == null) { errors.Add($"{oar.OarId}: source structure missing"); return; }

                    double thickness = oar.ParsedNestedThicknessMm ?? RCC_NESTED_RING_STEP_MM;
                    if (thickness <= 0) thickness = RCC_NESTED_RING_STEP_MM;

                    try
                    {
                        using (var tg = new TempGuard(_ss))
                        {
                            var prevShellByTarget = new Dictionary<string, SegmentVolume>(StringComparer.OrdinalIgnoreCase);
                            var targetStByTarget = new Dictionary<string, Structure>(StringComparer.OrdinalIgnoreCase);

                            foreach (var targetId in targetIds)
                            {
                                var targetSt = _ss.Structures.FirstOrDefault(s => !s.IsEmpty && string.Equals(s.Id, targetId, StringComparison.OrdinalIgnoreCase));
                                if (targetSt == null) { errors.Add($"{targetId}: target structure missing"); continue; }

                                var overlap = SafeBoolean(_ss, oarSt.SegmentVolume, targetSt.SegmentVolume, BoolOp.And,
                                    oarSt, targetSt, null, fb, $"RccNested_{oar.OarId}_{targetId}_Overlap", tg);
                                if (overlap == null) continue;

                                prevShellByTarget[targetId] = overlap;
                                targetStByTarget[targetId] = targetSt;
                            }

                            for (int level = 1; level <= RCC_NESTED_MAX_LEVELS && prevShellByTarget.Count > 0; level++)
                            {
                                var shellBase = SafeMargin(oarSt.SegmentVolume, -thickness * level);
                                var shellBaseSt = tg.Add(CreateTempFromSegment(_ss, shellBase, "zRCC_ShB"));

                                var levelPieces = new List<Structure>();
                                var nextShellByTarget = new Dictionary<string, SegmentVolume>(StringComparer.OrdinalIgnoreCase);

                                foreach (var kvp in prevShellByTarget)
                                {
                                    string targetId = kvp.Key;
                                    SegmentVolume prevShell = kvp.Value;
                                    var targetSt = targetStByTarget[targetId];

                                    var shell = SafeBoolean(_ss, shellBase, targetSt.SegmentVolume, BoolOp.And,
                                        shellBaseSt, targetSt, null, fb, $"RccNested_{oar.OarId}_{targetId}_Shell{level}", tg);

                                    SegmentVolume ringPiece;
                                    if (shell == null)
                                    {
                                        // OAR fully contracted out of this target - the
                                        // remaining piece from the previous level is the
                                        // last band; don't carry this target further.
                                        ringPiece = prevShell;
                                    }
                                    else
                                    {
                                        var prevShellSt = tg.Add(CreateTempFromSegment(_ss, prevShell, "zRCC_PrevSh"));
                                        var shellSt = tg.Add(CreateTempFromSegment(_ss, shell, "zRCC_Sh"));
                                        ringPiece = SafeBoolean(_ss, prevShell, shell, BoolOp.Sub,
                                            prevShellSt, shellSt, null, fb, $"RccNested_{oar.OarId}_{targetId}_Ring{level}", tg);
                                        nextShellByTarget[targetId] = shell;
                                    }

                                    if (ringPiece != null) levelPieces.Add(tg.Add(CreateTempFromSegment(_ss, ringPiece, "zRCC_RingPiece")));
                                }

                                if (levelPieces.Count > 0)
                                {
                                    var union = tg.Add(UnionManyToTemp(_ss, levelPieces, fb, "zRCC_ShellUn", $"NestedShell_{oar.OarId}_{level}", tg));
                                    if (union != null)
                                    {
                                        string shellId = RccNestedShellId(oar.OarId, level);
                                        var st = GetOrCreate(_ss, "CONTROL", shellId);
                                        EnsureRccHighRes(st);
                                        if (AssignSegmentSafely(st, union.SegmentVolume))
                                        {
                                            st.Color = level % 2 == 1 ? Colors.Gold : Colors.DarkGoldenrod;
                                            created.Add(shellId);
                                        }
                                        else
                                        {
                                            _ss.RemoveStructure(st);
                                            errors.Add($"{shellId}: empty result");
                                        }
                                    }
                                }

                                prevShellByTarget = nextShellByTarget;
                            }
                        }
                    }
                    catch (Exception ex) { errors.Add($"{oar.OarId}_in_ptv_hr#: {ex.Message}"); }
                }

            }
        }
    }
}
