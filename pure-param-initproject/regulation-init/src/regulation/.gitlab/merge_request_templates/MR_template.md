### Conformity: [PURE_RQM-XXXX] Summary

- [ ] Check the title of the MR: [PURE_RQM-XXXX] Summary
- [ ] Rebase the branch if needed (https://confluence.sbs-software.com/x/k7E3Lg)
- [ ] Elements modified are versionned (at least version is modified and startingValidityDate also if needed)
- [ ] The compatibiltyMatrix is updated (element version changed, startingValidityDate if needed)
- [ ] The packageMatrix (generic and by country) are updated if needed (refer to correct compatibilityMatrix output/regulation/common/ccm)
- [ ] Check modifications correspond to the expected one
- [ ] Check the e2e testData are updated (input and expected if available)
- [ ] Pipeline is manually launched with these modifications and is green
- [ ] Modification corresponds to process define in wiki pages (https://confluence.sbs-software.com/x/tguaOQ)

#### Forms
- [ ] Check the borders of each sheet
- [ ] Form
    - [ ] Check the name
    - [ ] Check the datatype
- [ ] Tables (sheet)
    - [ ] Check the name
    - [ ] Check the level of row
    - [ ] Check the rowNumbering
    - [ ] Check the columnNumbering
    - [ ] Check the labels
    - [ ] Check the datatypes if different of datatype defined on form side
    - [ ] Check the cell types (color: manual, computed, not required)
        - [ ] Check if the abstract row is marked as “yes” when the entire line is not required and no merge cells
    - [ ] Check the cell format according to the DPM (https://confluence.constellation.soprasteria.com/x/4Q77LQ)
    - [ ] Check the computation rules, computation rule order and cellSpecificities if needed
- [ ] Sheet "translation"
    - [ ] Check the label to see if it contains EN and FR columns
- [ ] All the yaml files are generated: main yaml, allocationRules, computationRules, datatypes, label-EN, label-FR, xref.
- [ ] Check the usertemplates
- [ ] Output by country are generated if needed
