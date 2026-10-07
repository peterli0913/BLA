#!/usr/bin/env python3
"""Apply the agreed English terminology / translation fixes to RA-139788_step3设备管线风险评估文件_改.docx."""

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

SRC = Path("/workspace/RA-139788_step3设备管线风险评估文件_改.docx")
OUT = Path("/workspace/RA-139788_step3设备管线风险评估文件_改_术语统一.docx")

TITLE_EN = "Risk Assessment of Equipment Piping for 139788 Step 3, TJ4 Site"

# (old, new). Applied in order to every paragraph (body, tables, TOC, headers, footers); all occurrences replaced.
RULES = [
    # ---------- cover, header, TOC, signature ----------
    ("TJ4 Factory 139788 step3 Equipment Material Conveying Pipeline Quality Inspection Scheme & Risk Assessment of Reporting Unsatisfactory Items",
     "Risk Assessment of Non-conformances Identified in the Material Transfer Piping Quality Inspection Protocol & Report for 139788 Step 3 Equipment, TJ4 Site"),
    ("Risk Assessment of Equipment Pipeline in TJ4 Factory Area 139788 step3", TITLE_EN),
    ("Cleaning Risk Assessment for TJ4 Site /", TITLE_EN + "/"),
    ("TJ4139788 step3设备管线风险评估", "TJ4厂区139788 step3设备管线风险评估"),
    ("Vice General Manager, QA", "Deputy General Manager, QA"),
    ("Table of Content/", "Table of Contents/"),
    ("3.1 Non-compliance Item 1: Assessment of the Non-BPE Piping of the 139788 step3 Preparation Solution System and Preparation System Piping;",
     "3.1 Non-conformance 1: Non-BPE (ISO) Piping in the 139788 Step 3 Preparative Buffer Preparation and Preparative Chromatography Systems"),
    ("3.2 Non-compliance Item 2: Situation Where Multiple Tees, Crosses and Other Positions of Piping Do Not Meet the 2D Requirement;",
     "3.2 Non-conformance 2: Tees, Crosses and Other Branches Not Meeting the 2D Dead-Leg Requirement"),
    ("3.3 Non-compliance Item 3: Situation Where Diaphragm Valve Installation Angle Does Not Meet Self-Draining Requirement.",
     "3.3 Non-conformance 3: Diaphragm Valves Not Installed at the Self-Draining Angle"),
    ("4.Conclusion /", "4. Conclusion /"),
    ("5.Revision History /", "5. Revision History /"),

    # ---------- 1. Objective / 2. Scope ----------
    ("Economic-Technological Development Area", "Economic and Technological Development Zone"),
    ("TJ4 site", "TJ4 Site"),
    ("The purpose of this risk assessment is to systematically evaluate the unsatisfied items in the quality inspection plan & report (SVP-TJ4-183.01-004) of the material conveying pipeline of 139788 step3 equipment.",
     "The purpose of this risk assessment is to systematically assess the non-conformances identified in the Material Transfer Piping Quality Inspection Protocol & Report (SVP-TJ4-183.01-004) for 139788 Step 3 equipment."),
    ("This appraisal document is applicable to the systematic appraisal of the unsatisfied items in the quality inspection plan & report (SVP-TJ4-183.01-004) of 139788 step3 material conveying pipeline of Asymchem Life Science (Tianjin) Co., Ltd (Tianjin) Co., Ltd. (No.265, South Street, Economic and Technological Development Zone, Binhai New Area, Tianjin, China). There are mainly three aspects that can not meet the requirements of the corresponding BPE standard, namely: 1.139788 step3 Evaluation of the preparation liquid preparation system and the pipeline of the preparation system as non-BPE pipeline; 2. The condition that multiple tees, tees and tank mouths of equipment and pipelines do not meet the requirements of 2D, and 3. The installation angle of diaphragm valve does not meet the requirements of self-exhaustion.",
     "This risk assessment applies to the systematic assessment of the non-conformances identified in the Material Transfer Piping Quality Inspection Protocol & Report (SVP-TJ4-183.01-004) for 139788 Step 3 at Asymchem Life Science (Tianjin) Co., Ltd. (No. 265, South Street, Economic and Technological Development Zone, Binhai New Area, Tianjin, China) (hereinafter referred to as the \"TJ4 Site\"). The non-conformances fall into three areas that do not meet the corresponding ASME BPE requirements: (1) non-BPE (ISO) piping used in the 139788 Step 3 preparative buffer preparation system and preparative chromatography system; (2) tees, crosses and vessel nozzles on equipment and piping not meeting the 2D dead-leg requirement (L/D ≤ 2); and (3) diaphragm valves not installed at the self-draining angle."),

    # ---------- 3. Risk assessment: introduction ----------
    ("The 139788step3 process pipeline is mainly used for system or solution transfer between devices in this step, and the process solution can be transferred in a closed way in the pipeline. At present, the pipeline is mainly made of 316L stainless steel with different specifications, and some of them are made of Hastelloy. Most of the process pipelines adopt quick-opening connection mode, including standard straight pipes with a length less than 2m. The pipeline system also includes corresponding standard quick-opening elbows, tees, four-way, etc., and valves, flowmeters, pumps, hoses and other designs are arranged in the pipeline to realize material transfer control and control residues to the greatest extent. Through the corresponding cleaning methods in and after the project batch, the pipeline can be effectively cleaned to avoid cross-contamination and control the load of microorganisms and endotoxin. Some of the hoses are lined with tetrafluoroethylene and the inner surface is smooth, which can meet the requirements of material compatibility and residue control.",
     "The 139788 Step 3 process piping is used mainly to transfer reaction mixtures and process solutions between equipment items in this step, allowing process solutions to be transferred in a closed system. The piping is mainly 316L stainless steel of various sizes, with some Hastelloy. Most process lines use sanitary clamp (tri-clamp) connections and include standard straight spools less than 2 m long. The piping system also includes standard clamp elbows, tees, crosses, etc., and valves, flowmeters, pumps and hoses are installed to control material transfer and minimize residues. The piping is effectively cleaned during and after project batches using the applicable cleaning methods, to prevent cross-contamination and control bioburden and endotoxin. Some hoses are PTFE-lined with a smooth internal surface, meeting material compatibility and residue control requirements."),
    ("In the 139788 step3 process piping, the national standard (ISO) piping adopts manual welding, while the BPE piping adopts automatic welding. All corresponding pipes undergo passivation work before use and piping validation work (Material Transfer Piping Quality Inspection Plan & Report (SVP-TJ4-183.01-004)). This document mainly conducts risk analysis on three aspects of non-compliance to ensure the corresponding risks are controllable, and formulates short-term and long-term measures",
     "In the 139788 Step 3 process piping, ISO-standard piping is manually welded, while BPE piping is orbitally (automatically) welded. All piping is passivated prior to use and qualified per the Material Transfer Piping Quality Inspection Protocol & Report (SVP-TJ4-183.01-004). This document performs an FMEA on the three non-conformances to ensure that the associated risks are controlled, and defines short-term and long-term measures"),
    ("（物料输送管路质量检测方案&报告（SVP-TJ4-183.01-004），", "（物料输送管路质量检测方案&报告（SVP-TJ4-183.01-004）），"),
    ("Description of FMEA scoring standard:/", "FMEA Scoring Criteria:/"),
    ("5 = Catastrophic (quality safety/regulatory non-compliance)",
     "5 = Catastrophic (impact on product quality/patient safety, or regulatory non-compliance)"),
    ("Probability of occurrence (p): 1 = rare occurrence (< 1%); 2= unlikely to happen (1%-5%); 3= Occasional (5-10%); 4= Frequent occurrence (10-30%); 5=Very Frequent occurrence (> 30%).",
     "Probability of Occurrence (P): 1 = Rare (<1%); 2 = Unlikely (1–5%); 3 = Occasional (5–10%); 4 = Frequent (10–30%); 5 = Very frequent (>30%)."),
    ("Detectability (D): 1 = almost certainly detectable (e.g. 100% endoscope+physical and chemical detection); 2= it can be detected with high probability (such as high-proportion sampling visual inspection+physical and chemical inspection); 3= There is a moderate chance of detection (such as visual inspection by random inspection); 5= Almost impossible to detect (e.g. only by final product inspection).",
     "Detectability (D): 1 = Almost certain detection (e.g., 100% borescope inspection plus physicochemical testing); 2 = High likelihood of detection (e.g., visual inspection at a high sampling rate plus physicochemical testing); 3 = Moderate likelihood of detection (e.g., visual inspection on a sampling basis); 5 = Almost impossible to detect (e.g., only by final product testing)."),

    # ---------- 3.1 ----------
    ("Specific Situation:", "Description:"),
    ("Specific Piping List:", "Piping List:"),
    ("Due to the long delivery lead time of BPE piping, temporary use was made for the early Demo batches of this 139788 step3 piping. After the subsequent BPE piping arrives, re-manufacturing and installation will be carried out. A comprehensive comparison and risk analysis between the existing ISO piping and BPE piping will be conducted to confirm whether it meets project production use.",
     "Due to the long lead time of BPE tubing, ISO piping was used on an interim basis for the early Demo batches of 139788 Step 3. After the BPE tubing arrives, the piping will be re-fabricated and re-installed using BPE tubing. This document compares the existing ISO piping with ASME BPE requirements and assesses whether it is suitable for production use."),
    ("Pipeline number/", "Line No./"),
    ("Specifications/", "Size/"),
    ("texture of wood/", "Material/"),
    ("executive standard/", "Applicable Standard/"),
    ("Peducer/", "Reducer/"),
    ("Acetonitrile pipeline/", "Acetonitrile line/"),
    ("A systematic assessment is conducted on core technical requirements such as material, welding, acid pickling passivation, and tolerance diameter, based on the ISO standard system requirements and referencing ASME BPE-2026 (Bioprocessing Equipment Standard, June 2026). The assessment follows ISO management requirements and covers the following core equipment and facility components: piping, pipe fittings, valves, pumps, and all surface treatment links in contact with the product. The assessment includes core requirements for material, welding, acid pickling passivation, tolerance diameter, and surface finish.",
     "A systematic assessment of the core technical requirements, including material, welding, pickling and passivation, and dimensional tolerances, was performed based on ISO tubing standard requirements, with reference to ASME BPE-2026 (Bioprocessing Equipment, June 2026). The assessment covers the following core equipment and facility components: piping, fittings, valves, pumps, and all product-contact surfaces and their surface treatment. The assessment includes the core requirements for material, welding, pickling and passivation, dimensional tolerances and surface finish."),
    ("评估以 ISO 管理要求，", ""),
    ("Material Requirements Assessment: Product contact surfaces are all based on 316L as the benchmark; non-contact surfaces can be downgraded to 304/304L. All the above related materials are provided with material certificates including material content. The requirements are basically consistent between ISO and ASME BPE-2026.",
     "Material Requirements Assessment: All product-contact surfaces use 316L as the baseline; non-product-contact surfaces may be 304/304L. Material test reports (MTRs) including chemical composition are provided for all of the above materials. Requirements are essentially equivalent between ISO and ASME BPE-2026."),
    ("Evaluation of Welding Requirements:The welding gas has been tested and 99.997% meets the argon shielding requirements. Automatic welding is prioritized, with manual welding as a supplement. The inspection ratio requires ≥20% endoscopic inspection for automatic welds and 100% endoscopic inspection for manual welds. The 2D dead leg rule (L/D ≤ 2:1) is applied, and product contact surfaces must be free from porosity, undercut, welding slag, and other defects. The requirements between ISO and ASME BPE-2026 are basically consistent, and the acceptance of ISO standard piping is based on the requirements of ASME BPE-2026. Welder qualifications are government-issued Special Operation Certificates",
     "Welding Requirements Assessment: Argon shielding gas was tested at 99.997% purity, meeting the shielding gas requirement. Orbital (automatic) welding is preferred; manual welding is used only where orbital welding is not practicable. Internal borescope examination: ≥ 20% of orbital (automatic) welds and 100% of manual welds. The 2D dead-leg rule (L/D ≤ 2) applies. Porosity, undercut, slag inclusions, etc. are not permitted on product-contact surfaces. Requirements are essentially equivalent between ISO and ASME BPE-2026, and ISO tubing is accepted against ASME BPE-2026 acceptance criteria. Welders hold government-issued Special Operations Certificates for welding"),
    ("This paper evaluates the inconsistency of pipeline system which adopts ISO standard pipes and is accepted according to ASME BPE. In welding, there are deviations between ISO pipes and BPE requirements in sulfur content, dimensional tolerance and end face quality, which may lead to defects such as insufficient fluidity, incomplete fusion, porosity, hot cracks or mismatching. As for the roughness of inner wall, ISO pipes usually only specify Ra value, while BPE has systematic requirements for surface smoothness of SFF1-SFF6, and high roughness will affect CIP/SIP effect. In view of the above risks, the following control measures are taken: 100% check the outer diameter, wall thickness and ovality of the pipe end, and control the wrong side; Re-evaluate the welding procedure with this batch of pipes, lock the welding parameters and record each joint; The sampling rate of automatic welding endoscope is 20%, and manual welding endoscope is 100%; The manufacturer provides the roughness test report for the inner wall roughness, and our company checks and accepts it according to BPE standard. Combined with cleaning verification and surface sampling monitoring, the welding defects and cleaning risks caused by the difference between ISO pipe and BPE standard can be effectively reduced. Therefore, the scheme is technically feasible and the differential risk is acceptable under controlled conditions",
     "The following assesses the gaps arising from accepting ISO tubing against ASME BPE criteria. For welding, ISO tubing differs from BPE requirements in sulfur content, dimensional tolerances and end-face quality, which may cause inconsistent orbital weld-pool flow, incomplete fusion, porosity, hot cracking or fit-up misalignment. For internal surface roughness, ISO tubing usually specifies only an Ra value, whereas ASME BPE defines surface finish designations SF1–SF6; higher roughness will reduce CIP/SIP effectiveness. To address these risks, the following controls are applied: 100% verification of tube-end OD, wall thickness and ovality to control fit-up misalignment; re-qualification of the welding procedure (WPS/PQR) using tubing from the same heat/lot, with fixed weld parameters and every weld recorded in the weld log/weld map; borescope inspection of 20% of orbital welds and 100% of manual welds; and verification of internal surface roughness against the supplier's roughness test report, accepted per ASME BPE. Combined with cleaning validation and surface (swab) sampling monitoring, these controls effectively reduce the welding-defect and cleaning risks arising from the differences between ISO tubing and the ASME BPE standard. Therefore, this approach is technically feasible, and the residual risk from these differences is acceptable under the stated controls"),
    ("SFF1–SFF6", "SF1–SF6"),
    ("Acid Pickling Passivation Requirements Assessment: There are acid pickling passivation plans, acid pickling passivation records, post-passivation test results, and polishing standards. Requirements are basically consistent between ISO and ASME BPE-2026.",
     "Pickling and Passivation Requirements Assessment: Pickling and passivation protocols, execution records, post-passivation test results and surface finish specifications are available. Requirements are essentially equivalent between ISO and ASME BPE-2026."),
    ("Tolerance and Diameter Requirements Assessment: Pipe connections use clamp quick connections, pipe slope satisfies >=1%, inner surface roughness in direct contact with product satisfies <=0.4um. Requirements are basically consistent between ISO and ASME BPE-2026. Pipe diameter, tolerance, wall thickness, etc., are selected in combination with different process points using pipe materials meeting ISO or ASME BPE-2026 process requirements; process risk is acceptable",
     "Tolerance and Diameter Requirements Assessment: Piping uses sanitary clamp (tri-clamp) connections; piping slope is ≥ 1%; product-contact surface roughness is Ra ≤ 0.4 µm. Requirements are essentially equivalent between ISO and ASME BPE-2026. Pipe diameter, tolerance, wall thickness, etc. are selected according to the requirements of each process point, using tubing that meets ISO or ASME BPE-2026 process requirements; the process risk is acceptable"),
    ("要求满足≤0.4 um", "要求满足Ra≤0.4 μm"),
    ("Short-term Measures: The overall management requirements for piping shall be consistent with BPE piping, supporting the production of the 139788 step3 Demo batch",
     "Short-term Measures: The ISO piping shall be managed under the same requirements as BPE piping, to support production of the 139788 Step 3 Demo batches"),
    ("Long-term Measures: Complete the re-design and installation validation of BPE piping before the 139788 step3 Pre-PPQ, subsequently serving as the dedicated piping for this project.",
     "Long-term Measures: Complete redesign, installation and qualification of BPE piping before 139788 Step 3 Pre-PPQ; this piping will then be dedicated to the project."),

    # ---------- FMEA tables (all three) ----------
    ("Targeted evaluation form FMEA/", "FMEA Worksheet/"),
    ("Failure mode/", "Failure Mode/"),
    ("Potential impact/", "Potential Effect/"),
    ("Before taking measures/", "Pre-mitigation/"),
    ("Control measure/", "Control Measures/"),
    ("After the measures/", "Post-mitigation (Residual)/"),
    ("Risk level/", "Risk Level/"),
    ("ISO/BPE hybrid welding produces staggered edges and incomplete fusion. /",
     "Misalignment and incomplete fusion at ISO-to-BPE welds/"),
    ("Affect the welding quality, which may lead to cleaning dead corners./",
     "Compromised weld quality; may create uncleanable crevices/"),
    ("100% recheck the pipe end size and control the wrong side; 100% endoscopic inspection of ISO-BPE butt weld; Re-evaluate the welding procedure /",
     "100% verification of tube-end dimensions to control misalignment; 100% borescope inspection of ISO-to-BPE butt welds; re-qualification of the welding procedure/"),
    ("The sulfur content of ISO pipeline is different from that of BPE standard, which affects the welding quality. /",
     "Sulfur content of ISO tubing differs from ASME BPE requirements, affecting weldability/"),
    ("It may lead to defects such as insufficient fluidity of molten pool, blowhole, hot crack and incomplete fusion, which will affect weld quality and cleaning risk./",
     "May cause inconsistent weld-pool flow, porosity, hot cracking and incomplete fusion, affecting weld quality and increasing cleaning risk/"),
    ("Check the sulfur content in the pipe material certificate, and retest by batch if necessary; 100% recheck the pipe end size and control the wrong side; Re-evaluate the welding process with this batch of pipes and lock the welding parameters; 100% endoscopic inspection of ISO-BPE butt weld; 20% of automatic welding endoscopes were sampled and 100% of manual welding endoscopes were inspected./",
     "Verify sulfur content against the tubing MTRs and retest per heat if necessary; 100% verification of tube-end dimensions to control misalignment; re-qualify the welding procedure using tubing from the same heat/lot and fix the weld parameters; 100% borescope inspection of ISO-to-BPE butt welds; borescope inspection of 20% of orbital welds and 100% of manual welds/"),
    ("The difference of inner wall roughness affects CIP/SIP effect./",
     "Differences in internal surface roughness reduce CIP/SIP effectiveness/"),
    ("May affect the cleaning effect/", "May reduce cleaning effectiveness/"),
    ("Medium risk/低风险", "Low risk/低风险"),
    ("The roughness test report of the inner wall is provided by the manufacturer, and our company checks and accepts it according to BPE standard, combined with cleaning verification and surface sampling monitoring./",
     "Internal surface roughness is verified against the supplier's roughness test report and accepted per ASME BPE, combined with cleaning validation and surface (swab) sampling monitoring./"),
    ("After systematic evaluation, all pipes are accepted according to ASME BPE-2026 requirements, forming a complete standard group in terms of welding procedure qualification, welder qualification and welding acceptance grade, providing an executable main framework, covering core management elements such as materials, welding methods and qualifications, basic requirements for passivation, drainage slope, dead angle rules and surface roughness, etc. The risk of difference is controllable, so it can be directly used as the basis for implementation, and the risk of project use is low",
     "After systematic evaluation, all tubing is accepted against ASME BPE-2026 requirements. ASME BPE-2026 provides a complete, executable framework covering welding procedure qualification, welder qualification and weld acceptance criteria, as well as core management elements such as materials, welding methods and qualifications, basic passivation requirements, drainage slope, dead-leg rules and surface roughness. The residual risk from the differences is controlled; ASME BPE-2026 is therefore applied directly as the acceptance basis, and the risk to the project is low"),

    # ---------- 3.2 ----------
    ("In clean fluid systems, for all non-continuous flow branches (tee branches, instrument interfaces, sampling ports, drainage ports, plugs, etc.), the ratio of the length L from the inner wall of the main pipeline to the blind end of the branch, to the branch's own inner diameter d must satisfy L/d <= 2 (i.e., the 2D requirement)",
     "In hygienic fluid systems, for all non-continuous-flow branches (tee branches, instrument connections, sample ports, drain ports, blind caps, etc.), the ratio of the length L, measured from the inner wall of the main line to the closed end of the branch, to the branch inner diameter D must satisfy L/D ≤ 2 (i.e., the 2D requirement)."),
    ("Due to the long delivery lead time of short-radius tees, T-type valves, zero static valves, GMP diaphragm valves, etc., this 139788 step3 piping was manufactured using general fittings, resulting in some components not meeting the 2D requirement, which was identified during the piping validation process; the specific list is shown in the table below:",
     "Due to the long lead times of short-outlet tees, T-body valves, zero-static (zero dead-leg) valves, sanitary diaphragm valves, etc., the 139788 Step 3 piping was fabricated with standard (general-purpose) fittings. As a result, some components do not meet the 2D requirement; these were identified during piping qualification and are listed in the table below:"),
    ("P&ID No. of component 部件号", "Component Tag No./部件号"),
    ("使用原则 (mm:mm)", "Acceptance Criterion/使用原则 (mm:mm)"),
    ("When L/d > 2, the turbulence of the main pipeline (CIP condition Reynolds number Re>=3000) cannot effectively entrain the fluid within the branch; a static retention dead zone forms inside the branch; the medium cannot circulate and renew, and cleaning fluid cannot effectively penetrate - this is the root cause of all risks",
     "When L/D > 2, turbulence in the main line (Reynolds number Re ≥ 3000 under CIP conditions) cannot effectively entrain the fluid in the branch; the branch becomes a stagnant dead leg, the medium cannot circulate and be renewed, and cleaning solution cannot penetrate effectively. This is the underlying cause of the risks below"),
    ("Microbial pollution and biofilm risk: the fluid in the dead zone is still, and the pipeline is completely exhausted at low points after use, and nitrogen purging further reduces the residue. After the pipeline is used, carry out the circulation cleaning of each port to ensure that all the paths are cleaned. At the same time, dismantle the tee and other parts that do not meet 2D requirements for visual inspection and confirmation, and increase the circulation cleaning and wiping sampling of the dead zone of the side port, which can reduce the risk to a low level.",
     "1. Microbial Contamination and Biofilm Risk: Fluid in the dead leg is stagnant. After use, the piping is fully drained at low points, and nitrogen purging further reduces residues. After use, each port is flow-through cleaned to ensure that all flow paths are cleaned. In addition, tees and other components not meeting the 2D requirement are disassembled for visual inspection, and flow-through cleaning and swab sampling of side-branch dead legs are added, reducing the risk to low."),
    ("微生物污染与生物膜风险：死区内流体静止", "1. 微生物污染与生物膜风险：死区内流体静止"),
    ("CIP Cleaning Effectiveness Failure Risk: The turbulence of the main pipeline cannot be transmitted into the dead zone; cleaning fluid can only form a weak laminar exchange at the dead zone inlet and cannot effectively scour the inner wall; residual materials and cleaning agents cannot be completely displaced, forming permanent residual. Based on this risk, increase flow-through cleaning of the side port dead zones and perform visual inspection confirmation after disassembly, which can reduce the risk to low risk.",
     "2. CIP Cleaning Effectiveness Failure Risk: Turbulence in the main line cannot be transmitted into the dead leg; cleaning solution forms only a weak laminar exchange at the dead-leg inlet and cannot effectively scour the inner wall; residual materials and cleaning agents cannot be fully displaced, resulting in persistent residues. To address this risk, flow-through cleaning of side-branch dead legs is added, and components are visually inspected after disassembly, reducing the risk to low."),
    ("Product Quality and Yield Risk: Product materials retained in the dead zone cannot be discharged with the main pipeline and remain after batch completion, causing fluctuations in critical quality attributes. Product retention loss leads to decreased product yield. Based on this risk characteristic, after project use, process solvent flushing and nitrogen purging are performed to avoid corresponding risks.",
     "3. Product Quality and Yield Risk: Product retained in the dead leg cannot be discharged with the main line and remains after batch completion, causing variability in critical quality attributes (CQAs). Product hold-up also reduces yield. To address this risk, the piping is rinsed with process solvent and purged with nitrogen after project use."),
    ("Equipment Corrosion and Maintenance Risk: Liquid retained in the dead zone for a long time destroys the hygienic grade piping inner surface finish (Ra≤0.4um), causing pitting corrosion and crevice corrosion; corrosion products further aggravate retention and pollution. For this risk, timely complete drainage is performed after material and cleaning transfer is completed, and since the corresponding materials in this project are non-corrosive, the risk is low.",
     "4. Equipment Corrosion and Maintenance Risk: Liquid retained in the dead leg for long periods degrades the internal surface finish of hygienic piping (Ra ≤ 0.4 µm), causing pitting and crevice corrosion; corrosion products further aggravate hold-up and contamination. To address this risk, the piping is drained promptly after material transfer and cleaning, and because the materials in this project are non-corrosive, the risk is low."),
    ("Short-term Measures: Execute according to the risk assessment plan to support Demo batch production. Timely complete drainage through process solvent flushing and nitrogen purging; for cleaning, flow-through cleaning of each port is performed and disassembly for visual inspection is carried out; corresponding risks are controllable.",
     "Short-term Measures: Execute per this risk assessment to support Demo batch production. Drain promptly by process solvent rinsing and nitrogen purging; for cleaning, flow-through clean each port and disassemble components for visual inspection; the associated risks are controlled."),
    ("Long-term Measures: Before the Pre-PPQ batch, carry out equipment modification and dimensional optimization: shorten over-specification branch pipes and strictly control L/d<=2; replace positions that cannot be shortened with zero dead-leg tees or clamp-style flush joints to structurally eliminate retention cavities. And conduct re-validation confirmation of piping and equipment.",
     "Long-term Measures: Before the Pre-PPQ batches, modify the equipment and optimize dimensions: shorten non-conforming branches to strictly meet L/D ≤ 2; where a branch cannot be shortened, replace it with zero-dead-leg valves/tees or flush-mounted clamp fittings to structurally eliminate hold-up cavities. Then re-qualify the piping and equipment."),
    ("Dead zone retention leads to microbial growth (water phase pipeline)/",
     "Dead-leg hold-up leads to microbial growth (aqueous lines)/"),
    ("Identify the list of water phase pipelines; Increase the side port dead zone circulation cleaning; Increase wiping sampling /",
     "Identify all aqueous lines; add flow-through cleaning of side-branch dead legs; add swab sampling/"),
    ("Disassembly does not meet the visual inspection of 2D parts; Increase physical and chemical detection of wiping sampling; Process solvent washing and nitrogen purging/",
     "Disassemble and visually inspect components not meeting the 2D requirement; add physicochemical testing of swab samples; process solvent rinse and nitrogen purging/"),
    ("Long-term detention leads to pipeline corrosion/", "Prolonged liquid hold-up causes corrosion/"),
    ("May corrode the pipeline/", "May corrode the piping/"),
    ("Drain it in time after use; Materials are non-corrosive; Periodic inspection/",
     "Drain promptly after use; materials are non-corrosive; periodic inspection/"),
    ("The existing equipment can support Demo batch production. After the Demo batch, re-design and modify non-conforming piping, then conduct validation after modification, ensuring this non-compliance item is resolved before Pre-PPQ.",
     "The existing equipment can support Demo batch production. After the Demo batches, the non-conforming piping will be redesigned, modified and re-qualified, ensuring this non-conformance is resolved before Pre-PPQ."),

    # ---------- 3.3 ----------
    ("This assessment is based on ASME BPE standards: clean pipeline short pipe slope >=1%, valves must follow the pipeline slope to achieve self-draining; China GMP: pipeline design and installation should avoid dead corners and blind pipes to prevent microbial growth and ensure complete system drainage. In this validation, the essential reason why diaphragm valves do not meet the requirement is that self-draining of weir-type hygienic grade diaphragm valves depends on gravity action. When the installation angle does not meet the standard (level without slope, insufficient slope, or reverse slope), a permanent accumulated liquid dead zone forms at the bottom of the valve cavity, which is the root cause of all risks. The specific valve installation angle follows the manufacturer's equipment manual and is measured and confirmed one by one during piping validation. The following is a list of diaphragm valves identified during validation that do not meet the standard:",
     "This assessment is based on ASME BPE (hygienic piping shall be sloped ≥ 1%, and valves shall follow the piping slope to be self-draining) and China GMP (piping design and installation shall avoid dead legs and blind ends to prevent microbial growth and ensure complete drainage of the system). The non-conformance identified during qualification arises because weir-type sanitary diaphragm valves drain by gravity only when installed at the manufacturer's specified drain angle. When the installation angle is not met (horizontal with no slope, insufficient slope or reverse slope), persistent liquid hold-up forms at the bottom of the valve cavity, which is the underlying cause of the risks below. The installation angle of each valve follows the manufacturer's manual and was measured and confirmed individually during piping qualification. Diaphragm valves identified during qualification as not meeting the requirement are listed below:"),
    ("Serial number/", "No./"),
    ("Valve number/", "Valve Tag No./"),
    ("Factory/厂家", "Manufacturer/厂家"),
    ("Target angle/", "Required Drain Angle/"),
    ("Current angle/", "As-installed Angle/"),
    ("The pipeline is blocked and cannot be adjusted./", "Obstructed by adjacent piping; cannot be adjusted./"),
    ("The valve hangs upside down and clings to the roof./", "Valve installed inverted, flush against the top plate./"),
    ("The valve is vertical and clings to the top plate./", "Valve installed vertically, flush against the top plate./"),
    ("Specific risks: Assess the risk that the liquid hold-up in the valve cavity becomes a hotbed for microbial colonization. The solvents of related valves involved in this production transfer are all organic solvents, which are not substrates for promoting microbial growth, and the risks can be eliminated;",
     "Microbial Colonization Risk: The risk of microbial colonization in liquid hold-up in the valve cavity was assessed. The solvents passing through the affected valves in this production are all organic solvents, which do not support microbial growth; this risk is considered negligible;"),
    ("Product Yield Loss Risk: After material piping transfer is completed, process solvent flushing and nitrogen purging are used; since the process piping itself has a slope, the loss risk is low;",
     "Product Yield Loss Risk: After material transfer, the piping is rinsed with process solvent and purged with nitrogen; because the process piping is itself sloped, the risk of loss is low;"),
    ("Risk Assessment of Residual Finished Products/Materials Unable to Be Discharged from Accumulated Liquid Prolonged Immersion of the Diaphragm (EPDM/PTFE material), causing swelling, hardening, cracking, and shortened service life. Since this usage period is short, only Demo batch production is executed, and immediately after completion corresponding cleaning of equipment and pipelines is performed; there is no corresponding risk;",
     "Diaphragm Degradation Risk: Prolonged immersion of EPDM/PTFE diaphragms in retained product or material that cannot be drained may cause swelling, hardening and cracking, shortening diaphragm life. Because the period of use is short (Demo batch production only) and the equipment and piping are cleaned immediately afterwards, this risk is considered negligible;"),
    ("Short-term Measures: Execute according to the risk assessment plan to support Demo batch production. Timely complete drainage through process solvent flushing and nitrogen purging; for cleaning, disassemble non-conforming diaphragm valves for visual inspection; corresponding risks are controllable.",
     "Short-term Measures: Execute per this risk assessment to support Demo batch production. Drain promptly by process solvent rinsing and nitrogen purging; for cleaning, disassemble non-conforming diaphragm valves for visual inspection; the associated risks are controlled."),
    ("Long-term Measures: Before the Pre-PPQ batch, carry out equipment and piping modification to structurally eliminate the diaphragm valve installation angle problem. And conduct re-validation confirmation of piping and equipment.",
     "Long-term Measures: Before the Pre-PPQ batches, modify the equipment and piping to structurally eliminate the diaphragm valve installation angle issue. Then re-qualify the piping and equipment."),
    ("Clear material classification; Dismantle the unsatisfactory diaphragm valve for visual inspection; Increase wipe sampling (physicochemical/microbial)/",
     "Classify lines as aqueous or organic; disassemble non-conforming diaphragm valves for visual inspection; add swab sampling (physicochemical/microbial)/"),
    ("Long-term immersion swelling, hardening and cracking of diaphragm/",
     "Diaphragm swelling, hardening or cracking due to prolonged immersion/"),
    ("Clean immediately after the end; Pre-PPQ transformation/", "clean immediately after use; modification before Pre-PPQ/"),

    # ---------- 4. Conclusion / 5. Revision history ----------
    ("Through the above risk analysis, methods such as flushing, nitrogen purging, disassembly cleaning confirmation can reduce the risk of non-compliance items in the 139788 step3 equipment/piping validation to low risk, supporting Demo batch production; after the Demo batch, re-design and modify non-conforming piping, then conduct validation after modification, ensuring this non-compliance item is resolved before Pre-PPQ.",
     "Based on the above risk analysis, rinsing, nitrogen purging, and disassembly with post-cleaning visual verification reduce the risk of the non-conformances identified in 139788 Step 3 equipment/piping qualification to low, supporting Demo batch production. After the Demo batches, the non-conforming piping will be redesigned, modified and re-qualified, ensuring these non-conformances are resolved before Pre-PPQ."),
    ("Date of the last signatory/", "Date of last signature/"),
    ("New version/", "Initial issue/"),
    ("In this risk assessment, a systematic risk assessment is carried out for the unsatisfied items in the quality inspection plan & report (SVP-TJ4-183.01-004) of the material conveying pipeline of 139788 step3 equipment.",
     "This risk assessment systematically assesses the non-conformances identified in the Material Transfer Piping Quality Inspection Protocol & Report (SVP-TJ4-183.01-004) for 139788 Step 3 equipment."),

    # ---------- global: L/D notation (EN, CN and the 2D table) ----------
    ("L/d", "L/D"),
    ("d=d0-2e2", "D=d0-2e2"),
    ("支路自身内径d", "支路自身内径D"),
]


def replace_in_paragraph(p, old, new):
    ts = [t for t in p.iter(qn("w:t"))]
    if not ts:
        return 0
    text = "".join(t.text or "" for t in ts)
    if old not in text:
        return 0
    starts, pos = [], 0
    for t in ts:
        starts.append(pos)
        pos += len(t.text or "")

    def locate(idx):
        for i in range(len(ts) - 1, -1, -1):
            if starts[i] <= idx:
                return i, idx - starts[i]
        return 0, idx

    hits, i = [], text.find(old)
    while i != -1:
        hits.append(i)
        i = text.find(old, i + len(old))
    for h in reversed(hits):
        si, so = locate(h)
        ei, eo = locate(h + len(old) - 1)
        eo += 1
        first = ts[si].text or ""
        if si == ei:
            ts[si].text = first[:so] + new + first[eo:]
        else:
            last = ts[ei].text or ""
            ts[si].text = first[:so] + new
            for k in range(si + 1, ei):
                ts[k].text = ""
            ts[ei].text = last[eo:]
        for k in (si, ei):
            ts[k].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return len(hits)


def all_roots(doc):
    roots = [doc.element.body]
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/header") or rel.reltype.endswith("/footer"):
            roots.append(rel.target_part.element)
    return roots


def main():
    doc = Document(SRC)
    paras = [p for root in all_roots(doc) for p in root.iter(qn("w:p"))]
    missed = []
    for old, new in RULES:
        n = sum(replace_in_paragraph(p, old, new) for p in paras)
        if n == 0:
            missed.append(old[:70])
    doc.save(OUT)
    print(OUT)
    print("rules:", len(RULES), "missed:", len(missed))
    for m in missed:
        print("  MISSED:", m)


if __name__ == "__main__":
    main()
