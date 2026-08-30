const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageOrientation, Header, Footer, PageNumber, TabStopType,
} = require('docx');

const ACCENT = '0F5257';
const GREY = '5A6A6B';
const RULE = 'C4CCCB';
const SOFT = 'EDF2F1';

// A4 content width in DXA: 11906 - 2*1134 ≈ 9638
const W = 9350;

const p = (text, opts = {}) => new Paragraph({
  spacing: { before: opts.before ?? 0, after: opts.after ?? 120, line: 276 },
  alignment: opts.align,
  indent: opts.indent,
  children: Array.isArray(text) ? text : [new TextRun({ text, size: opts.size ?? 20, bold: opts.bold, italics: opts.italics, color: opts.color, font: 'Calibri' })],
  border: opts.border,
});

const run = (t, o = {}) => new TextRun({ text: t, size: o.size ?? 20, bold: o.bold, italics: o.italics, color: o.color, font: 'Calibri' });

const h1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 360, after: 160 },
  children: [new TextRun({ text: t, size: 28, bold: true, color: ACCENT, font: 'Calibri' })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: RULE, space: 6 } },
});

const h2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 240, after: 100 },
  children: [new TextRun({ text: t, size: 23, bold: true, font: 'Calibri' })],
});

// numbered clause: bold number + text
const cl = (num, text, opts = {}) => new Paragraph({
  spacing: { before: 60, after: 100, line: 276 },
  indent: { left: 640, hanging: 640 },
  children: [
    new TextRun({ text: num + '\t', bold: true, size: 22, font: 'Calibri' }),
    ...(Array.isArray(text) ? text : [new TextRun({ text, size: 22, font: 'Calibri' })]),
  ],
  tabStops: [{ type: TabStopType.LEFT, position: 640 }],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: 'dash', level: 0 },
  spacing: { before: 40, after: 60, line: 276 },
  children: Array.isArray(text) ? text : [new TextRun({ text, size: 22, font: 'Calibri' })],
});

const noBorders = {
  top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
  insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
};

const tblBorders = {
  top: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: RULE },
  insideVertical: { style: BorderStyle.NONE },
};

function table(rows, widths, opts = {}) {
  const header = opts.header;
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: opts.borders ?? tblBorders,
    rows: rows.map((cells, ri) => new TableRow({
      tableHeader: header && ri === 0,
      children: cells.map((c, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA },
        margins: { top: 90, bottom: 90, left: 110, right: 110 },
        shading: (header && ri === 0) ? { type: ShadingType.CLEAR, fill: SOFT } : undefined,
        children: [new Paragraph({
          spacing: { before: 0, after: 0, line: 260 },
          alignment: opts.rightCols && opts.rightCols.includes(ci) ? AlignmentType.RIGHT : undefined,
          children: Array.isArray(c) ? c : [new TextRun({
            text: String(c), size: 21, font: 'Calibri',
            bold: (header && ri === 0) || undefined,
            color: (header && ri === 0) ? GREY : undefined,
          })],
        })],
      })),
    })),
  });
}

const fillRow = (label) => new Paragraph({
  spacing: { before: 200, after: 40 },
  tabStops: [{ type: TabStopType.LEFT, position: 3200 }],
  children: [
    new TextRun({ text: label + '\t', size: 22, font: 'Calibri' }),
    new TextRun({ text: ' '.repeat(60), size: 22, font: 'Calibri', underline: {} }),
  ],
});

const callout = (label, text) => new Table({
  columnWidths: [W],
  width: { size: W, type: WidthType.DXA },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    left: { style: BorderStyle.SINGLE, size: 18, color: ACCENT },
    right: { style: BorderStyle.SINGLE, size: 4, color: RULE },
    insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE },
  },
  rows: [new TableRow({
    children: [new TableCell({
      width: { size: W, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: SOFT },
      margins: { top: 160, bottom: 160, left: 200, right: 200 },
      children: [
        new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: label.toUpperCase(), size: 17, bold: true, color: GREY, font: 'Calibri', characterSpacing: 20 })] }),
        new Paragraph({ spacing: { after: 0, line: 276 }, children: Array.isArray(text) ? text : [new TextRun({ text, size: 22, font: 'Calibri' })] }),
      ],
    })],
  })],
});

const F = (t) => new TextRun({ text: t, size: 22, font: 'Calibri', color: 'A0392B', bold: true }); // fill-in field
const B = (t) => new TextRun({ text: t, size: 22, font: 'Calibri', bold: true });
const T = (t) => new TextRun({ text: t, size: 22, font: 'Calibri' });
const I = (t) => new TextRun({ text: t, size: 22, font: 'Calibri', italics: true });

const children = [];

// ---------- Title block ----------
children.push(new Paragraph({
  spacing: { after: 60 },
  children: [new TextRun({ text: 'CASUAL EMPLOYMENT AGREEMENT', size: 36, bold: true, font: 'Calibri' })],
}));
children.push(new Paragraph({
  spacing: { after: 200 },
  children: [new TextRun({ text: 'Builder — QBCC Nominee Supervisor', size: 26, color: ACCENT, font: 'Calibri' })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } },
}));

children.push(callout('Draft for review', [
  T('Complete every '), F('[bracketed]'), T(' field, have a Queensland construction and employment lawyer review the document, then sign. Background and reasoning are in the accompanying structuring guidance note.'),
]));

children.push(p('', { after: 240 }));
children.push(p([T('THIS AGREEMENT is made on '), F('[date]')]));
children.push(p([B('BETWEEN: '), B('FIRST LIGHT CIVIL PTY LTD'), T(' ACN '), F('[ACN]'), T(' ABN '), F('[ABN]'), T(' of '), F('[registered office]'), T(', Caboolture, Queensland ('), B('"FLC"'), T(')')]));
children.push(p([B('AND: '), B('ALISTAIR COLEMAN'), T(' of '), F('[address]'), T(' ('), B('"the Employee"'), T(')')], { after: 240 }));

// ---------- 1 ----------
children.push(h1('1. The job'));
children.push(cl('1.1', [T('FLC employs the Employee as a '), B('casual employee'), T(' in the position of '), B('Builder — QBCC Nominee Supervisor'), T(', starting '), F('[date]'), T('.')]));
children.push(cl('1.2', [T('The Employee is appointed as FLC\'s '), B('nominee'), T(' under section 42B of the '), I('Queensland Building and Construction Commission Act 1991'), T(' (Qld) ('), B('"QBCC Act"'), T(') for FLC\'s licence no. '), F('[FLC licence no.]'), T(', class '), F('[class]'), T('.')]));
children.push(cl('1.3', [T('The Employee reports to '), F('[Director name]'), T(', Director.')]));
children.push(cl('1.4', [T('Work is on FLC\'s retaining wall projects across '), B('Brisbane and South East Queensland'), T('. FLC\'s project locations are not limited to any suburb or council area.')]));

// ---------- 2 ----------
children.push(h1('2. Casual employment — what that means'));
children.push(cl('2.1', [T('The employment is casual. There is '), B('no firm advance commitment'), T(' by FLC to continuing and indefinite work, and none by the Employee to perform it. Accordingly:')]));
children.push(bullet('FLC need not offer any Engagement, and the Employee need not accept one;'));
children.push(bullet('there is no guaranteed roster, pattern, or number of hours or jobs; and'));
children.push(bullet([T('the Employee is paid a '), B('25% casual loading'), T(' in recognition of that.')]));
children.push(cl('2.2', [T('An '), B('"Engagement"'), T(' is a discrete period of work on a specific project, offered by FLC and accepted by the Employee using '), B('Schedule 1'), T('.')]));
children.push(cl('2.3', [T('The Employee acknowledges receipt of the '), B('Fair Work Information Statement'), T(' and the '), B('Casual Employment Information Statement'), T('. FLC will give a further copy of the Casual Employment Information Statement at 12 months.')]));
children.push(cl('2.4', [T('The Employee may notify FLC under the Employee Choice Pathway (Division 4A, Part 2-2, '), I('Fair Work Act 2009'), T(' (Cth)) if he believes he no longer meets the casual definition. FLC will respond as required.')]));

// ---------- 3 ----------
children.push(h1('3. Scope of FLC\'s work'));
children.push(cl('3.1', [T('FLC builds only retaining walls that do '), B('not'), T(' require engineering certification under a local law or the Building Regulation ('), B('"Exempt Walls"'), T('), and FLC\'s QBCC licence class is limited accordingly.')]));
children.push(cl('3.2', 'Walls falling outside that scope are dealt with under clause 4.4.'));

// ---------- 4 ----------
children.push(h1('4. The Employee\'s duties'));
children.push(cl('4.1', [B('Core obligation. '), T('To '), B('adequately supervise'), T(' all retaining wall building work carried out under FLC\'s QBCC licence, within the meaning of section 43A of the QBCC Act.')]));
children.push(cl('4.2', [B('Pre-start scope assessment — the Employee\'s primary duty. '), T('Before FLC quotes any wall, the Employee must assess and confirm in writing whether it is an Exempt Wall, having regard to:')]));
children.push(p('', { after: 60 }));
children.push(table([
  ['(a)', 'Finished height, including retained soil'],
  ['(b)', 'Any surcharge load above or behind — driveway, slab, pool, shed, vehicle access, building'],
  ['(c)', 'Clear distance to any building or other retaining wall'],
  ['(d)', 'Whether the wall is tiered, stepped or terraced, or interacts with an existing wall'],
  ['(e)', 'Whether it forms part of a pool barrier'],
  ['(f)', 'Soil and site conditions — reactive clay, uncontrolled fill, drainage'],
  ['(g)', 'Heritage, easement or other site constraint'],
], [700, W - 700]));
children.push(p('', { after: 120 }));
children.push(p([T('He must also confirm the work is '), B('within the scope of FLC\'s licence class'), T(', and tell FLC in writing where it is not.')], { after: 160 }));

children.push(cl('4.3', [B('Supervision. '), T('The Employee will:')]));
children.push(bullet('establish a system of supervision proportionate to the size and complexity of the work;'));
children.push(bullet([T('attend site and inspect at the nominated hold points — '), B('at minimum excavation and founding material, wall construction, drainage and filter media, and backfill and compaction'), T(' — recording each visit using '), B('Schedule 2'), T(';')]));
children.push(bullet([T('verify the work as constructed conforms to the manufacturer\'s specifications, '), B('AS 4678 '), I('Earth-retaining structures'), T(' and the National Construction Code, and direct rectification where it does not;')]));
children.push(bullet('direct FLC\'s employees and subcontractors on site as to how the building work is performed; and'));
children.push(bullet('perform all duties of a nominee under the QBCC Act.'));

children.push(cl('4.4', [B('If a wall is not an Exempt Wall. '), T('The Employee must notify FLC in writing immediately, and:')]));
children.push(bullet([T('FLC must '), B('not commence or continue'), T(' the work until an RPEQ structural engineer is engaged, any required design certification (Form 15) and building approval obtained, and FLC has confirmed the work is within its licence class;')]));
children.push(bullet('if the work proceeds, the Employee\'s duties extend to reviewing the RPEQ-certified design and approval conditions and coordinating the engineer\'s and certifier\'s inspections and certificates; and'));
children.push(bullet('the additional time is paid at the Hourly Rate, and clause 6.6 applies if Budgeted Hours are exceeded.'));

children.push(cl('4.5', [B('Limits on the role. '), T('The parties acknowledge:')]));
children.push(bullet([T('structural design certification is the function of an '), B('RPEQ engineer'), T(', not the Employee;')]));
children.push(bullet([T('building approval and final certification are functions of the '), B('building certifier'), T(', not the Employee;')]));
children.push(bullet([T('FLC will not describe the Employee, in any contract, quote, marketing material or communication, as certifying or "signing off" a retaining wall. '), B('He supervises; an RPEQ certifies'), T('; and')]));
children.push(bullet([T('the Employee\'s role is '), B('supervisory and technical'), T('. He is not engaged to perform on-site trade work, operate plant, or work on the tools, and FLC will not direct him to. Changing this requires a written variation and a review of pay, as it may alter which industrial instrument applies.')]));

children.push(cl('4.6', [B('Records. '), T('The Employee must complete a scope assessment for every wall and a Schedule 2 record for every site visit. FLC must keep them for '), B('7 years'), T('. These records are the evidence of adequate supervision if QBCC audits FLC.')]));
children.push(cl('4.7', [B('No licence lending. '), T('Neither party will enter into or continue any arrangement under which the Employee is held out as nominee for work he does not in fact adequately supervise. FLC must give him genuine and unimpeded access to sites, documents and people, and must not allow work to proceed contrary to his written direction under clause 4.2 or 4.4.')]));

// ---------- 5 ----------
children.push(h1('5. Licences and disclosure'));
children.push(cl('5.1', [T('The Employee warrants he holds, and will maintain, a current QBCC '), B('contractor\'s licence no. 76041, class Builder \u2014 Low Rise, expiring 16 March 2027'), T(', together with any further class required to match FLC\'s company licence class, free of any condition preventing him acting as nominee.')]));
children.push(cl('5.1A', [T('The Employee must '), B('renew licence 76041 before 16 March 2027'), T(' and give FLC evidence of renewal. FLC will diarise the date. The parties acknowledge that if the licence lapses, FLC has no nominee, must notify QBCC within 14 days, and commits an offence if it remains without a nominee for 28 days or more (clause 12.3).')]));
children.push(cl('5.2', [T('He must notify FLC '), B('within 2 business days'), T(' if his licence is suspended, cancelled, conditioned or expires; if he becomes bankrupt or an excluded individual; or if anything else affects his eligibility.')]));
children.push(cl('5.3', 'He has disclosed every other licensee for which he acts as nominee, warrants he can adequately perform the role for all of them, and must notify FLC before accepting another appointment.'));
children.push(cl('5.4', 'FLC must notify him immediately of anything affecting FLC\'s licence.'));
children.push(cl('5.5', [T('The parties acknowledge the Employee may be an '), B('"influential person"'), T(' for FLC, so each party\'s licensing history may affect the other\'s licence.')]));

// ---------- 6 ----------
children.push(h1('6. Pay'));
children.push(callout('In plain terms', [B('The Employee is paid 5% of each job.'), T(' The hourly rate is the legal mechanism that ensures he is paid for his time; the completion payment tops him up to the full 5%.')]));
children.push(p('', { after: 160 }));
children.push(cl('6.1', [B('Hourly Rate. '), B('$'), F('[70]'), B(' per hour'), T(', inclusive of the 25% casual loading, for every hour worked.')]));
children.push(cl('6.2', [B('Minimum rates. '), T('This Agreement is made on the basis that the Employee is '), B('not covered by a modern award'), T(', being engaged for licensed technical judgement and supervision rather than trade work, so the applicable minimum is the National Minimum Wage. The Hourly Rate exceeds it. '), B('If a modern award does apply, that award prevails'), T(' — FLC will pay not less than the applicable award rate including any loading, allowance and minimum engagement, and will make good any shortfall.')]));
children.push(cl('6.3', [B('Job Value Allowance. '), T('Each project carries a '), B('Job Value Allowance of 5% of the Contract Value'), T(', being the target total cost of the Employee\'s wages '), B('and'), T(' superannuation for that project.')]));
children.push(p('', { after: 60 }));
children.push(table([
  ['Term', 'Meaning'],
  [[B('Contract Value')], [T('Total payable by FLC\'s client for the retaining wall works, '), B('excluding GST'), T(', and '), F('[including / excluding]'), T(' approved variations')]],
  [[B('Wages Component')], 'Job Value Allowance ÷ 1.12 (superannuation at 12%) = 4.464% of Contract Value'],
  [[B('Budgeted Hours')], 'Wages Component ÷ Hourly Rate — recorded in Schedule 1 before work starts'],
], [2400, W - 2400], { header: true }));
children.push(p('', { after: 160 }));
children.push(cl('6.4', [B('Completion payment. '), T('In the pay period after practical completion, FLC compares total wages paid for the project against the Wages Component. If wages paid are '), B('less'), T(', FLC pays the difference as a completion payment. It is wages — PAYG is withheld and superannuation applies.')]));
children.push(cl('6.5', [B('No clawback. '), T('If wages paid '), B('exceed'), T(' the Wages Component, no adjustment is made and '), B('no amount is recoverable from the Employee'), T('. FLC will not deduct or set off any amount against wages already earned.')]));
children.push(cl('6.6', [B('Scope control. '), T('The Employee must get FLC\'s '), B('written approval before exceeding Budgeted Hours'), T('. Failing to do so is a performance matter and is '), B('not'), T(' a basis for FLC to refuse or reduce payment for hours actually worked.')]));
children.push(cl('6.7', [B('The Hourly Rate is the floor. '), T('The Employee is entitled to the Hourly Rate for every hour worked regardless of the Budgeted Hours, any completion payment, or whether the project proceeds, is cancelled, varied or unprofitable.')]));
children.push(cl('6.8', 'The casual loading is paid in lieu of paid annual leave, personal/carer\'s leave, compassionate leave, notice of termination, redundancy pay, and public holidays not worked.'));
children.push(cl('6.9', [B('Payment. '), F('[Weekly / fortnightly]'), T(' by electronic transfer, with a payslip within one working day. FLC withholds PAYG as required.')]));
children.push(p('', { after: 80 }));
children.push(callout('Example — a $20,000 wall', [
  T('Job Value Allowance $1,000. Wages Component $892.86, superannuation $107.14, Budgeted Hours 12.75. The Employee works 4.5 hours: paid $315 as he goes, then a '), B('$577.86 completion payment'), T('. Total received: '), B('$1,000 — exactly 5%.'),
]));

// ---------- 7 ----------
children.push(h1('7. Superannuation'));
children.push(cl('7.1', [T('FLC pays superannuation at the rate required by law — currently '), B('12%'), T(' of ordinary time earnings — to a complying fund the Employee nominates, or his stapled fund.')]));
children.push(cl('7.2', [T('Contributions are paid within the time required by law, including the '), B('payday superannuation'), T(' requirements from 1 July 2026 (within 7 business days of paying wages).')]));
children.push(cl('7.3', 'Superannuation applies to the Hourly Rate (including casual loading) and to any completion payment, to the extent each is ordinary time earnings.'));

// ---------- 8 ----------
children.push(h1('8. Hours, travel and expenses'));
children.push(cl('8.1', 'There are no set hours. Hours for each Engagement are as agreed in Schedule 1 and as reasonably required to discharge clause 4.'));
children.push(cl('8.2', [T('The Employee records all hours and submits them '), F('[weekly / on completion of each Engagement]'), T('.')]));
children.push(cl('8.3', [B('Travel time. '), T('Paid hours are time on site plus associated document review, assessment and record-keeping. '), B('Travel to and from sites is not paid time'), T(' — the Hourly Rate is set on the basis that it compensates for travel within Brisbane and South East Queensland, and clause 8.4 covers vehicle running costs.')]));
children.push(cl('8.4', [B('Vehicle. '), F('[The Employee uses his own vehicle and FLC pays $[rate] per kilometre / FLC provides a vehicle.]')]));
children.push(cl('8.5', [T('Clause 8.3 does not reduce total pay for any period below the minimum payable under the '), I('Fair Work Act 2009'), T(' (Cth) or any applicable award across all hours that constitute work. If an award applies and prescribes a fares and travel allowance, FLC pays it in addition.')]));
children.push(cl('8.6', 'FLC reimburses reasonable pre-approved out-of-pocket expenses on production of receipts.'));
children.push(cl('8.7', [F('[The Employee provides his own measuring and testing equipment. FLC provides all personal protective equipment.]')]));
children.push(cl('8.8', 'Where a modern award applies and prescribes a minimum engagement, the Employee is paid that minimum for each engagement. A single engagement may cover more than one project, with hours apportioned between them for clause 6.'));

// ---------- 9 ----------
children.push(h1('9. Entitlements'));
children.push(cl('9.1', 'As a casual, the Employee has the National Employment Standards entitlements that apply to casuals — including unpaid carer\'s leave, unpaid compassionate leave, family and domestic violence leave, community service leave, and after 12 months of regular and systematic employment, unpaid parental leave and the right to request flexible working.'));
children.push(cl('9.2', [T('FLC will register with '), B('QLeave'), T(' and lodge the returns recording the Employee\'s service for portable long service leave.')]));
children.push(cl('9.3', [T('FLC will maintain a '), B('WorkCover Queensland'), T(' policy covering the Employee.')]));

// ---------- 10 ----------
children.push(h1('10. Work health and safety'));
children.push(cl('10.1', [T('The Employee must comply with the '), I('Work Health and Safety Act 2011'), T(' (Qld), applicable codes of practice, and FLC\'s WHS policies and site rules.')]));
children.push(cl('10.2', 'He must not allow work to proceed where it would create a risk to health or safety, and must report all incidents, injuries and hazards immediately.'));
children.push(cl('10.3', 'FLC will provide a safe system of work and will not pressure the Employee to approve, permit or overlook work that does not comply with the approved design or applicable standards.'));

// ---------- 11 ----------
children.push(h1('11. Confidentiality and conflicts'));
children.push(cl('11.1', 'The Employee must keep FLC\'s client lists, pricing, quotations and methods confidential, during and after employment, except as required by law or as necessary to discharge his statutory duties.'));
children.push(cl('11.2', 'He must disclose any actual or potential conflict of interest, including any competing retaining wall or earthworks business he operates or works for.'));
children.push(cl('11.3', 'Nothing here prevents a disclosure to QBCC, the Fair Work Ombudsman or any regulator where the disclosure is required or protected by law.'));
children.push(cl('11.4', 'Records and documents created in the course of an Engagement are FLC\'s property, subject to the Employee\'s right to keep copies relating to his own statutory obligations.'));

// ---------- 12 ----------
children.push(h1('12. Ending the employment'));
children.push(cl('12.1', [T('Either party may end an Engagement, or the employment, on '), F('[1 day\'s]'), T(' notice, or immediately by agreement. No notice payment or redundancy pay applies — the casual loading is paid in lieu.')]));
children.push(cl('12.2', 'FLC may terminate immediately for serious misconduct, or if the Employee ceases to hold the licence required by clause 5.1.'));
children.push(cl('12.3', [B('QBCC notification. '), T('The parties acknowledge:')]));
children.push(bullet([T('if the Employee ceases to act as nominee, '), B('QBCC must be notified within 14 days'), T(';')]));
children.push(bullet([T('FLC must not carry out or undertake building work while it has no nominee, and commits an offence if it is without one for '), B('28 days or more'), T('; and')]));
children.push(bullet('each party will sign whatever QBCC forms are needed to give effect to this.'));
children.push(cl('12.4', [T('The Employee must give FLC '), B('['), F('4'), B('] weeks\' notice'), T(' before resigning as nominee (as distinct from declining Engagements), so FLC can appoint a replacement inside the 28-day window. This clause survives termination.')]));
children.push(cl('12.5', 'On termination the Employee returns all FLC property and records.'));

// ---------- 13 ----------
children.push(h1('13. General'));
children.push(cl('13.1', [T('This Agreement is governed by the laws of '), B('Queensland'), T('.')]));
children.push(cl('13.2', 'It is the entire agreement between the parties and replaces all prior agreements and understandings.'));
children.push(cl('13.3', 'It may only be varied in writing signed by both parties, except that FLC may increase the Hourly Rate unilaterally.'));
children.push(cl('13.4', [B('Statutory entitlements prevail. '), T('Nothing in this Agreement excludes, modifies or reduces any entitlement under the '), I('Fair Work Act 2009'), T(' (Cth), the National Employment Standards, any applicable modern award, or any other law. To the extent of any inconsistency, that law prevails.')]));
children.push(cl('13.5', 'If any clause is invalid or unenforceable it is severed and the rest continues.'));

// ---------- Signing ----------
children.push(h1('Signed'));
children.push(p([B('FIRST LIGHT CIVIL PTY LTD'), T(' ACN '), F('[ACN]')], { before: 120 }));
children.push(fillRow('Director signature'));
children.push(fillRow('Name'));
children.push(fillRow('Date'));
children.push(p([B('ALISTAIR COLEMAN')], { before: 320 }));
children.push(fillRow('Signature'));
children.push(fillRow('Date'));

// ================= SCHEDULE 1 =================
children.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 60 }, children: [new TextRun({ text: 'SCHEDULE 1', size: 32, bold: true, font: 'Calibri' })] }));
children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Engagement Confirmation', size: 26, color: ACCENT, font: 'Calibri' })], border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } } }));
children.push(p([I('One per wall. Complete the scope assessment before FLC quotes.')], { after: 200 }));

children.push(table([
  ['Project', ''],
  ['Site address', ''],
  ['FLC job number', ''],
  ['Client', ''],
  ['Description of work', ''],
  ['Expected start / completion', ''],
], [3000, W - 3000], { header: true }));

children.push(h2('Scope assessment — clause 4.2'));
children.push(table([
  ['Check', 'Answer'],
  ['Finished height incl. retained soil', '________ mm'],
  ['Surcharge load above or behind?', 'Yes  /  No'],
  ['Clear distance to nearest building or wall', '________ m'],
  ['Tiered, stepped, or near an existing wall?', 'Yes  /  No'],
  ['Part of a pool barrier?', 'Yes  /  No'],
  ['Soil or drainage concerns?', 'Yes  /  No'],
  ['Heritage, easement or other constraint?', 'Yes  /  No'],
  [[B('Is this an Exempt Wall?')], [B('Yes  /  No')]],
  [[B('Within FLC\'s licence class?')], [B('Yes  /  No')]],
], [W - 2600, 2600], { header: true }));
children.push(p('', { after: 100 }));
children.push(callout('Stop condition', [B('If either bolded answer is No, clause 4.4 applies. Do not quote or start until an RPEQ engineer is engaged and the licence class is confirmed.')]));
children.push(fillRow('Nominee signature'));
children.push(fillRow('Date'));

children.push(h2('Commercials'));
children.push(table([
  ['Contract Value (ex GST)', '$'],
  [[B('Job Value Allowance (5%)')], [B('$')]],
  ['Wages Component (÷ 1.12)', '$'],
  ['Superannuation (12%)', '$'],
  ['Hourly Rate', '$'],
  [[B('Budgeted Hours')], [B('________ h')]],
], [W - 3000, 3000]));

children.push(h2('Offer and acceptance'));
children.push(p([T('FLC offers this Engagement on the terms of the Casual Employment Agreement dated '), F('[date]'), T('.')]));
children.push(fillRow('FLC'));
children.push(fillRow('Date'));
children.push(p('The Employee accepts / declines, and acknowledges this Engagement carries no commitment to further work.', { before: 200 }));
children.push(fillRow('Employee'));
children.push(fillRow('Date'));

children.push(h2('Hours worked'));
children.push(table([
  ['Date', 'Activity', 'Hours'],
  ['', '', ''], ['', '', ''], ['', '', ''], ['', '', ''],
  ['', [B('Total')], ''],
], [1800, W - 3300, 1500], { header: true }));

children.push(h2('Completion reconciliation — clause 6.4'));
children.push(table([
  ['Wages paid for this project', '$'],
  ['Less Wages Component', '$'],
  [[B('Completion payment due (if positive)')], [B('$')]],
  ['Superannuation on total wages (12%)', '$'],
  ['Budgeted Hours exceeded?', 'Yes / No — written approval: Yes / No / N/A'],
], [W - 3600, 3600]));

// ================= SCHEDULE 2 =================
children.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 60 }, children: [new TextRun({ text: 'SCHEDULE 2', size: 32, bold: true, font: 'Calibri' })] }));
children.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: 'Site Supervision Record', size: 26, color: ACCENT, font: 'Calibri' })], border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: ACCENT, space: 8 } } }));
children.push(p([I('One per site visit. Keep for 7 years. This is FLC\'s evidence of adequate supervision under section 43A of the QBCC Act.')], { after: 200 }));

children.push(table([
  ['Project', ''],
  ['Job no.', ''],
  ['Date', ''],
  ['On site', ''],
  ['Off site', ''],
], [3000, W - 3000]));

children.push(h2('Inspected this visit'));
['Scope assessment confirmed on site — still an Exempt Wall',
 'Excavation / founding material',
 'Wall construction / block fill / core fill',
 'Drainage, filter media, geofabric',
 'Backfill and compaction',
 'Completion / final inspection',
 'Other: ______________________________'].forEach(x => children.push(bullet(x)));

children.push(h2('Observations'));
children.push(table([[''], [''], ['']], [W]));

children.push(h2('Checks'));
children.push(table([
  [[B('Conforms to manufacturer\'s specifications, AS 4678 and the NCC?')], [B('Yes  /  No')]],
  [[B('Any change on site affecting the Exempt Wall assessment?')], [B('Yes  /  No')]],
], [W - 2600, 2600]));
children.push(p([I('If Yes to the second question — stop work and apply clause 4.4.')], { after: 160 }));

children.push(h2('Directions given / rectification required'));
children.push(table([[''], ['']], [W]));

children.push(h2('Sign off'));
children.push(table([
  ['Rectification verified complete?', 'Yes / No / N/A — date: ____________'],
], [W - 4200, 4200]));
children.push(fillRow('Nominee signature'));
children.push(fillRow('QBCC licence no.'));

const doc = new Document({
  creator: 'First Light Civil Pty Ltd',
  title: 'Casual Employment Agreement — Builder, QBCC Nominee Supervisor',
  description: 'Casual employment agreement between First Light Civil Pty Ltd and Alistair Coleman',
  numbering: {
    config: [{
      reference: 'dash',
      levels: [{
        level: 0, format: 'bullet', text: '–', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1000, hanging: 300 } } },
      }],
    }],
  },
  styles: {
    default: {
      document: { run: { font: 'Calibri', size: 22 } },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838, orientation: PageOrientation.PORTRAIT },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          spacing: { after: 200 },
          tabStops: [{ type: TabStopType.RIGHT, position: W }],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: RULE, space: 4 } },
          children: [new TextRun({ text: 'First Light Civil Pty Ltd  ·  Casual Employment Agreement\tDRAFT', size: 17, color: GREY, font: 'Calibri' })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], size: 17, color: GREY, font: 'Calibri' })],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(process.argv[2], b);
  console.log('written', process.argv[2], b.length, 'bytes');
});
