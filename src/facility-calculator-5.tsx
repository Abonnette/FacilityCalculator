// @ts-nocheck

import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';

export default function FacilityCalculator() {
  const [step, setStep] = useState(0);
  const [linkInputs, setLinkInputs] = useState({});
  const [structureNames, setStructureNames] = useState({
    type1: 'Tangent',
    type2: 'Medium Running Angle',
    type3: 'Angle Deadend',
    type4: 'Full Tension Deadend',
    type5: 'N/A'
  });
  const [customMode, setCustomMode] = useState({
    type1: false,
    type2: false,
    type3: false,
    type4: false,
    type5: false
  });
  const [voltage, setVoltage] = useState('');
  const [rowWidth, setRowWidth] = useState('');
  const [numCircuits, setNumCircuits] = useState('1');
  const [conductor1Spec, setConductor1Spec] = useState('');
  const [conductor2Spec, setConductor2Spec] = useState('');
  const [shieldWire1Spec, setShieldWire1Spec] = useState('');
  const [shieldWire2Spec, setShieldWire2Spec] = useState('');
  const [structureFamily, setStructureFamily] = useState('');
  const [structureCosts, setStructureCosts] = useState({
    type1Material: 0,
    type1Labor: 0,
    type2Material: 0,
    type2Labor: 0,
    type3Material: 0,
    type3Labor: 0,
    type4Material: 0,
    type4Labor: 0,
    type5Material: 0,
    type5Labor: 0
  });
  const [links, setLinks] = useState([
    { id: Date.now(), link: 'A', length: '1', structureType1: '1', structureType2: '', structureType3: '', structureType4: '', structureType5: '' },
    { id: Date.now() + 1, link: 'B', length: '2', structureType1: '', structureType2: '1', structureType3: '', structureType4: '', structureType5: '' },
    { id: Date.now() + 2, link: 'C', length: '3', structureType1: '', structureType2: '', structureType3: '1', structureType4: '', structureType5: '' },
    { id: Date.now() + 3, link: 'D', length: '4', structureType1: '', structureType2: '', structureType3: '', structureType4: '1', structureType5: '' }
  ]);
  
  // Initialize default routes with link combinations
  const defaultLinkIds = {
    A: Date.now(),
    B: Date.now() + 1,
    C: Date.now() + 2,
    D: Date.now() + 3
  };
  
  const [routes, setRoutes] = useState([
    { id: Date.now() + 100, name: '1', linkIds: [defaultLinkIds.A, defaultLinkIds.B] },
    { id: Date.now() + 101, name: '2', linkIds: [defaultLinkIds.B, defaultLinkIds.C] },
    { id: Date.now() + 102, name: '3', linkIds: [defaultLinkIds.C, defaultLinkIds.D] },
    { id: Date.now() + 103, name: '4', linkIds: [defaultLinkIds.A, defaultLinkIds.D] },
    { id: Date.now() + 104, name: '5', linkIds: [defaultLinkIds.B, defaultLinkIds.D] },
    { id: Date.now() + 105, name: '6', linkIds: [defaultLinkIds.A, defaultLinkIds.C] }
  ]);
  const [subCategories, setSubCategories] = useState({});
  const [subCategoryMethods, setSubCategoryMethods] = useState({});
  const [inputs, setInputs] = useState({});
  const [total, setTotal] = useState(null);
  const [routeLandAcquisition, setRouteLandAcquisition] = useState({
    [Date.now() + 100]: 1000000,
    [Date.now() + 101]: 1000000,
    [Date.now() + 102]: 1000000,
    [Date.now() + 103]: 1000000,
    [Date.now() + 104]: 1000000,
    [Date.now() + 105]: 1000000
  }); // Store land acquisition cost per route ID
  const [substationCost, setSubstationCost] = useState('2000000');
  const [contingency, setContingency] = useState(15);
  const [tempContingency, setTempContingency] = useState('15');
  const [contingencyChanged, setContingencyChanged] = useState(false);
  const [escalationRate, setEscalationRate] = useState(3);
  const [escalationYears, setEscalationYears] = useState(2);
  const [escalationMonths, setEscalationMonths] = useState(0);
  const [rates, setRates] = useState({ 
    tax: 8.25, 
    stores: 5.0, 
    overhead: 8.275,
    rowAgent: 15000, 
    surveying: 0,
    aerialSurvey: 0,
    soilBorings: 0,
    engineering: 0,
    clearing: 0,
    environmental: 0,
    damages: 0,
    constructionMgmt: 0,
    permitting: 0,
    wireLabor: 0,
    conductor2Labor: 0,
    shieldWire1Labor: 0,
    shieldWire2Labor: 0,
    conductor1Material: 0,
    conductor2Material: 0,
    shieldWire1Material: 0,
    shieldWire2Material: 0
  });
  const [tempSubCat, setTempSubCat] = useState('');
  const [editingCat, setEditingCat] = useState(null);

  const categories = [
    { key: 'rightOfWay', label: 'Right-of-Way and Land Acquisition', fixed: ['Land Acquisition', 'ROW Agent'] },
    { key: 'engineeringUtility', label: 'Engineering and Design (Utility)' },
    { key: 'engineeringContract', label: 'Engineering and Design (Contract)', fixed: ['Land Surveying and SUE', 'Aerial Survey', 'Soil Borings', 'Engineering and Design'] },
    { key: 'procurement', label: 'Procurement of Material and Equipment (including stores)', fixed: ['Structure Material', '1st Circuit Conductor Material', '2nd Circuit Conductor Material', '1st Circuit Shield Wire Material', '2nd Circuit Shield Wire Material', 'Material and Sales Tax', 'Purchases and Stores'] },
    { key: 'transmissionUtility', label: 'Construction of Facilities (Utility)' },
    { key: 'transmissionContract', label: 'Construction of Facilities (Contract)', fixed: ['Clearing', 'Environmental Controls', 'Damages', 'Construction Management', 'Permitting', 'Structure Labor', '1st Circuit Conductor Labor', '2nd Circuit Conductor Labor', '1st Circuit Shield Wire Labor', '2nd Circuit Shield Wire Labor', 'Construction Overhead'] },
    { key: 'other', label: 'Other (all costs not included in the above categories)' },
    { key: 'substationUtility', label: 'Estimated Substation Facilities Cost' }
  ];

  React.useEffect(() => {
    const fixed = {};
    const methods = {};
    categories.forEach(c => { 
      if (c.fixed) {
        fixed[c.key] = [...c.fixed];
        // Set default methods for fixed categories
        c.fixed.forEach((subcat, idx) => {
          const key = `${c.key}_${idx}`;
          if (c.key === 'rightOfWay' && idx === 1) methods[key] = '$/mile';
          else if (c.key === 'engineeringContract') methods[key] = '$/mile';
          else if (c.key === 'procurement' && idx >= 1 && idx <= 4) methods[key] = '$/mile'; // All conductor and shield wire materials
          else if (c.key === 'procurement' && (idx === 5 || idx === 6)) methods[key] = '%';
          else if (c.key === 'transmissionContract' && idx >= 0 && idx <= 4) methods[key] = '$/mile'; // Clearing through Permitting
          else if (c.key === 'transmissionContract' && idx === 5) methods[key] = 'fixed'; // Structure Labor
          else if (c.key === 'transmissionContract' && idx >= 6 && idx <= 9) methods[key] = '$/mile'; // All conductor/shield wire labor
          else if (c.key === 'transmissionContract' && idx === 10) methods[key] = '%';
          else methods[key] = 'fixed';
        });
      }
    });
    setSubCategories(fixed);
    setSubCategoryMethods(methods);
  }, []);

  // Sync tempContingency when step changes to 4
  useEffect(() => {
    if (step === 4) {
      setTempContingency(contingency.toString());
      setContingencyChanged(false);
    }
  }, [step]);

  const addLink = () => setLinks([...links, { id: Date.now(), link: '', length: '', structureType1: '', structureType2: '', structureType3: '', structureType4: '', structureType5: '' }]);
  const updateLink = (id, field, val) => setLinks(links.map(l => l.id === id ? { ...l, [field]: val } : l));
  const removeLink = (id) => setLinks(links.filter(l => l.id !== id));

  const downloadTemplate = () => {
    const headers = ['Link', 'Length (Miles)', ...[1, 2, 3, 4, 5].filter(n => structureNames[`type${n}`] !== 'N/A').map(n => structureNames[`type${n}`])];
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    const colWidths = [{ wch: 20 }, { wch: 15 }];
    [1, 2, 3, 4, 5].filter(n => structureNames[`type${n}`] !== 'N/A').forEach(() => colWidths.push({ wch: 20 }));
    ws['!cols'] = colWidths;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Links');
    XLSX.writeFile(wb, 'Links_Template.xlsx');
  };

  const uploadLinks = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
        setLinks(data.slice(1).filter(r => r[0]).map((r, i) => ({ id: Date.now() + i, link: r[0] || '', length: r[1] || '', structureType1: r[2] || '', structureType2: r[3] || '', structureType3: r[4] || '', structureType4: r[5] || '', structureType5: r[6] || '' })));
        e.target.value = '';
      } catch { alert('Error reading file.'); }
    };
    reader.readAsArrayBuffer(file);
  };

  const addSubCat = (k) => {
    if (!tempSubCat.trim()) return;
    const newSubs = { ...subCategories, [k]: [...(subCategories[k] || []), tempSubCat.trim()] };
    const newIdx = (subCategories[k] || []).length;
    const newMethods = { ...subCategoryMethods, [`${k}_${newIdx}`]: 'fixed' };
    setSubCategories(newSubs);
    setSubCategoryMethods(newMethods);
    setTempSubCat('');
  };

  const removeSubCat = (k, i) => {
    const cat = categories.find(c => c.key === k);
    if (cat.fixed && i < cat.fixed.length) return;
    const newSubs = { ...subCategories };
    newSubs[k].splice(i, 1);
    if (newSubs[k].length === 0) delete newSubs[k];
    
    // Remove method and reindex remaining methods
    const newMethods = { ...subCategoryMethods };
    delete newMethods[`${k}_${i}`];
    // Reindex methods after the removed one
    for (let j = i + 1; j < subCategories[k].length; j++) {
      const oldKey = `${k}_${j}`;
      const newKey = `${k}_${j - 1}`;
      if (newMethods[oldKey]) {
        newMethods[newKey] = newMethods[oldKey];
        delete newMethods[oldKey];
      }
    }
    
    setSubCategories(newSubs);
    setSubCategoryMethods(newMethods);
  };

  const initCosts = () => {
    const newInputs = {};
    Object.keys(subCategories).forEach(mk => {
      subCategories[mk].forEach((sc, si) => {
        links.forEach(l => {
          const key = `${mk}_${si}_${l.id}`;
          const len = parseFloat(l.length) || 0;
          const methodKey = `${mk}_${si}`;
          const method = subCategoryMethods[methodKey] || 'fixed';
          
          // Calculate structure material costs (procurement category, index 0)
          if (mk === 'procurement' && si === 0) {
            const st1 = parseFloat(l.structureType1) || 0;
            const st2 = parseFloat(l.structureType2) || 0;
            const st3 = parseFloat(l.structureType3) || 0;
            const st4 = parseFloat(l.structureType4) || 0;
            const st5 = parseFloat(l.structureType5) || 0;
            const totalMaterial = 
              (st1 * structureCosts.type1Material) +
              (st2 * structureCosts.type2Material) +
              (st3 * structureCosts.type3Material) +
              (st4 * structureCosts.type4Material) +
              (st5 * structureCosts.type5Material);
            newInputs[key] = totalMaterial.toFixed(2);
          }
          // Calculate structure labor costs (transmissionContract category, index 5)
          else if (mk === 'transmissionContract' && si === 5) {
            const st1 = parseFloat(l.structureType1) || 0;
            const st2 = parseFloat(l.structureType2) || 0;
            const st3 = parseFloat(l.structureType3) || 0;
            const st4 = parseFloat(l.structureType4) || 0;
            const st5 = parseFloat(l.structureType5) || 0;
            const totalLabor = 
              (st1 * structureCosts.type1Labor) +
              (st2 * structureCosts.type2Labor) +
              (st3 * structureCosts.type3Labor) +
              (st4 * structureCosts.type4Labor) +
              (st5 * structureCosts.type5Labor);
            newInputs[key] = totalLabor.toFixed(2);
          }
          // Calculate tax (procurement category, index 5)
          else if (mk === 'procurement' && si === 5) {
            const sm = parseFloat(newInputs[`procurement_0_${l.id}`]) || 0;
            const c1m = parseFloat(newInputs[`procurement_1_${l.id}`]) || 0;
            const c2m = parseFloat(newInputs[`procurement_2_${l.id}`]) || 0;
            const sw1m = parseFloat(newInputs[`procurement_3_${l.id}`]) || 0;
            const sw2m = parseFloat(newInputs[`procurement_4_${l.id}`]) || 0;
            const totalMaterial = sm + c1m + c2m + sw1m + sw2m;
            newInputs[key] = (totalMaterial * rates.tax / 100).toFixed(2);
          }
          // Calculate stores (procurement category, index 6)
          else if (mk === 'procurement' && si === 6) {
            const sm = parseFloat(newInputs[`procurement_0_${l.id}`]) || 0;
            const c1m = parseFloat(newInputs[`procurement_1_${l.id}`]) || 0;
            const c2m = parseFloat(newInputs[`procurement_2_${l.id}`]) || 0;
            const sw1m = parseFloat(newInputs[`procurement_3_${l.id}`]) || 0;
            const sw2m = parseFloat(newInputs[`procurement_4_${l.id}`]) || 0;
            const totalMaterial = sm + c1m + c2m + sw1m + sw2m;
            newInputs[key] = (totalMaterial * rates.stores / 100).toFixed(2);
          }
          // Initialize based on method for items with default rates
          else if (method === '$/mile' && rates[getMethodRateKey(mk, si)]) {
            newInputs[key] = (len * rates[getMethodRateKey(mk, si)]).toFixed(2);
          }
          else if (mk === 'transmissionContract' && si === 10) {
            let base = 0;
            for (let i = 0; i < 10; i++) {
              const k = `${mk}_${i}_${l.id}`;
              base += parseFloat(newInputs[k]) || 0;
            }
            newInputs[key] = (base * rates.overhead / 100).toFixed(2);
          }
          else newInputs[key] = '';
        });
      });
    });
    setInputs(newInputs);
    setStep(3);
  };

  const getMethodRateKey = (mk, si) => {
    if (mk === 'rightOfWay' && si === 1) return 'rowAgent';
    if (mk === 'engineeringContract' && si === 0) return 'surveying';
    if (mk === 'engineeringContract' && si === 1) return 'aerialSurvey';
    if (mk === 'engineeringContract' && si === 2) return 'soilBorings';
    if (mk === 'engineeringContract' && si === 3) return 'engineering';
    if (mk === 'procurement' && si === 1) return 'conductor1Material';
    if (mk === 'procurement' && si === 2) return 'conductor2Material';
    if (mk === 'procurement' && si === 3) return 'shieldWire1Material';
    if (mk === 'procurement' && si === 4) return 'shieldWire2Material';
    if (mk === 'transmissionContract' && si === 0) return 'clearing';
    if (mk === 'transmissionContract' && si === 1) return 'environmental';
    if (mk === 'transmissionContract' && si === 2) return 'damages';
    if (mk === 'transmissionContract' && si === 3) return 'constructionMgmt';
    if (mk === 'transmissionContract' && si === 4) return 'permitting';
    if (mk === 'transmissionContract' && si === 6) return 'wireLabor';
    if (mk === 'transmissionContract' && si === 7) return 'conductor2Labor';
    if (mk === 'transmissionContract' && si === 8) return 'shieldWire1Labor';
    if (mk === 'transmissionContract' && si === 9) return 'shieldWire2Labor';
    return null;
  };

  const updateInput = (key, val) => {
    const parts = key.split('_');
    if (parts[0] === 'rightOfWay' && parts[1] === '1') return;
    if (parts[0] === 'engineeringUtility' && (parts[1] === '0' || parts[1] === '1' || parts[1] === '2')) return;
    if (parts[0] === 'procurement' && parts[1] === '0') return; // Structure Material is auto-calculated
    if (parts[0] === 'transmissionContract' && (parts[1] === '0' || parts[1] === '1' || parts[1] === '2' || parts[1] === '3' || parts[1] === '4' || parts[1] === '5' || parts[1] === '6' || parts[1] === '7' || parts[1] === '8' || parts[1] === '9')) return;
    
    const newInputs = { ...inputs, [key]: val };
    if (parts[0] === 'procurement' && (parts[1] === '1' || parts[1] === '2' || parts[1] === '3' || parts[1] === '4')) {
      links.forEach(l => {
        const sm = parseFloat(newInputs[`procurement_0_${l.id}`]) || 0;
        const c1m = parseFloat(newInputs[`procurement_1_${l.id}`]) || 0;
        const c2m = parseFloat(newInputs[`procurement_2_${l.id}`]) || 0;
        const sw1m = parseFloat(newInputs[`procurement_3_${l.id}`]) || 0;
        const sw2m = parseFloat(newInputs[`procurement_4_${l.id}`]) || 0;
        const totalMaterial = sm + c1m + c2m + sw1m + sw2m;
        newInputs[`procurement_5_${l.id}`] = (totalMaterial * rates.tax / 100).toFixed(2);
        newInputs[`procurement_6_${l.id}`] = (totalMaterial * rates.stores / 100).toFixed(2);
      });
    }
    setInputs(newInputs);
    setTotal(Object.values(newInputs).reduce((a, v) => a + (isNaN(parseFloat(v)) || v === '' ? 0 : Math.ceil(parseFloat(v) / 1000) * 1000), 0));
  };

  const updateRate = (type, val) => {
    const newRates = { ...rates, [type]: val };
    setRates(newRates);
    const newInputs = { ...inputs };
    if (type === 'tax') links.forEach(l => { const b = (parseFloat(newInputs[`procurement_0_${l.id}`]) || 0) + (parseFloat(newInputs[`procurement_1_${l.id}`]) || 0); newInputs[`procurement_2_${l.id}`] = (b * val / 100).toFixed(2); });
    if (type === 'stores') links.forEach(l => { const b = (parseFloat(newInputs[`procurement_0_${l.id}`]) || 0) + (parseFloat(newInputs[`procurement_1_${l.id}`]) || 0); newInputs[`procurement_3_${l.id}`] = (b * val / 100).toFixed(2); });
    if (type === 'overhead') links.forEach(l => { 
      let base = 0;
      for (let i = 0; i < 10; i++) base += parseFloat(newInputs[`transmissionContract_${i}_${l.id}`]) || 0;
      newInputs[`transmissionContract_10_${l.id}`] = (base * val / 100).toFixed(2);
    });
    if (type === 'rowAgent') links.forEach(l => newInputs[`rightOfWay_1_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'surveying') links.forEach(l => newInputs[`engineeringUtility_0_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'aerialSurvey') links.forEach(l => newInputs[`engineeringUtility_1_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'soilBorings') links.forEach(l => newInputs[`engineeringUtility_2_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'engineering') links.forEach(l => newInputs[`engineeringUtility_3_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'clearing') links.forEach(l => newInputs[`transmissionContract_0_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'environmental') links.forEach(l => newInputs[`transmissionContract_1_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'damages') links.forEach(l => newInputs[`transmissionContract_2_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'constructionMgmt') links.forEach(l => newInputs[`transmissionContract_3_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'permitting') links.forEach(l => newInputs[`transmissionContract_4_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'wireLabor') links.forEach(l => newInputs[`transmissionContract_6_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'conductor2Labor') links.forEach(l => newInputs[`transmissionContract_7_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'shieldWire1Labor') links.forEach(l => newInputs[`transmissionContract_8_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'shieldWire2Labor') links.forEach(l => newInputs[`transmissionContract_9_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'conductor1Material') links.forEach(l => newInputs[`procurement_1_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'conductor2Material') links.forEach(l => newInputs[`procurement_2_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'shieldWire1Material') links.forEach(l => newInputs[`procurement_3_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    if (type === 'shieldWire2Material') links.forEach(l => newInputs[`procurement_4_${l.id}`] = ((parseFloat(l.length) || 0) * val).toFixed(2));
    setInputs(newInputs);
    setTotal(Object.values(newInputs).reduce((a, v) => a + (isNaN(parseFloat(v)) || v === '' ? 0 : Math.ceil(parseFloat(v) / 1000) * 1000), 0));
  };

  const round = (v) => { const n = parseFloat(v); return isNaN(n) || v === '' ? 0 : Math.ceil(n / 1000) * 1000; };
  const subTotal = (mk, si) => links.reduce((a, l) => a + round(inputs[`${mk}_${si}_${l.id}`]), 0);
  const catTotal = (mk) => !subCategories[mk] ? 0 : subCategories[mk].reduce((a, sc, i) => a + subTotal(mk, i), 0);
  const transTotal = () => ['rightOfWay', 'engineeringUtility', 'engineeringContract', 'procurement', 'transmissionUtility', 'transmissionContract', 'other'].reduce((a, k) => a + catTotal(k), 0);

  const exportExcel = () => {
    const data = [['Category', 'Sub-Category', 'Amount']];
    categories.forEach(mc => {
      if (subCategories[mc.key]) {
        subCategories[mc.key].forEach((sc, i) => data.push([mc.label, sc, subTotal(mc.key, i)]));
        data.push([mc.label + ' Total', '', catTotal(mc.key)], []);
      }
    });
    data.push(['Estimated Total Transmission Cost', '', transTotal()], ['Estimated Substation Facilities Cost', '', catTotal('substationUtility')], ['Estimated Total Project Costs', '', total]);
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 50 }, { wch: 40 }, { wch: 20 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Facility Costs');
    XLSX.writeFile(wb, `Facility_Costs_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportLinkSummary = () => {
    // Build header row with link identifiers
    const headers = ['Category'];
    links.forEach(link => {
      headers.push(`Link ${link.link}`);
    });
    
    const data = [headers];
    
    // Calculate escalation and contingency factors
    const totalYears = escalationYears + (escalationMonths / 12);
    const escalationFactor = Math.pow(1 + (escalationRate / 100), totalYears);
    const contingencyFactor = 1 + (contingency / 100);
    
    // Calculate costs for all links
    const allLinkCosts = links.map(link => {
      const linkCategoryTotals = {};
      
      categories.slice(0, 7).forEach(mc => {
        if (subCategories[mc.key]) {
          let categoryTotal = 0;
          subCategories[mc.key].forEach((sc, si) => {
            const key = `${mc.key}_${si}_${link.id}`;
            categoryTotal += round(inputs[key]);
          });
          linkCategoryTotals[mc.key] = categoryTotal;
        }
      });
      
      // Apply escalation and contingency to each category, then round up to nearest 1,000
      const linkCategoryTotalsAdjusted = {};
      categories.slice(0, 7).forEach(mc => {
        const baseValue = linkCategoryTotals[mc.key] || 0;
        
        // Apply escalation
        const afterEscalation = baseValue * escalationFactor;
        
        // Apply contingency
        const afterContingency = afterEscalation * contingencyFactor;
        
        // Round final category value up to nearest 1,000
        linkCategoryTotalsAdjusted[mc.key] = Math.ceil(afterContingency / 1000) * 1000;
      });
      
      // Calculate link total from rounded category values
      const linkTotal = Object.values(linkCategoryTotalsAdjusted).reduce((sum, val) => sum + val, 0);
      
      return { linkCategoryTotalsAdjusted, linkTotal };
    });
    
    // Add category rows (adjusted values)
    categories.slice(0, 7).forEach(mc => {
      const row = [mc.label];
      allLinkCosts.forEach(linkCost => {
        row.push(linkCost.linkCategoryTotalsAdjusted[mc.key] || 0);
      });
      data.push(row);
    });
    
    // Add Estimated Total Cost row
    const totalRow = ['Estimated Total Cost'];
    allLinkCosts.forEach(linkCost => {
      totalRow.push(linkCost.linkTotal);
    });
    data.push(totalRow);
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // Set column widths
    const colWidths = [{ wch: 50 }]; // Category column
    links.forEach(() => colWidths.push({ wch: 15 })); // Link columns
    ws['!cols'] = colWidths;
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Link Cost Summary');
    XLSX.writeFile(wb, `Link_Cost_Summary_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportStructureCriteria = () => {
    // Current criteria sheet
    const data = [
      ['Voltage', voltage],
      ['Structure Family', structureFamily],
      ['Number of Circuits', numCircuits],
      [],
      ['Structure Type', 'Name', 'Material Cost', 'Labor Cost']
    ];
    
    for (let i = 1; i <= 5; i++) {
      const name = structureNames[`type${i}`];
      const material = structureCosts[`type${i}Material`] || 0;
      const labor = structureCosts[`type${i}Labor`] || 0;
      data.push([`Structure #${i}`, name, material, labor]);
    }
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 20 }, { wch: 30 }, { wch: 15 }, { wch: 15 }];
    
    // Defaults sheet with all scenarios
    const defaultsData = [
      ['Default Settings for All Voltage/Structure Family Scenarios'],
      [],
      ['69 kV - Steel Monopole'],
      ['Number of Circuits', '2 (default)'],
      ['ROW Width', '70 ft'],
      ['Structure Type', 'Name', 'Material Cost', 'Labor Cost'],
      ['Structure #1', 'Tangent', 46000, 22000],
      ['Structure #2', 'Small Running Angle', 65000, 87000],
      ['Structure #3', 'Medium Running Angle', 75000, 45000],
      ['Structure #4', 'Angle Deadend', 90000, 72000],
      ['Structure #5', 'Full Tension Deadend', 119000, 145000],
      [],
      ['138 kV - Steel Monopole'],
      ['Number of Circuits', '2 (default)'],
      ['ROW Width', '70 ft'],
      ['Structure Type', 'Name', 'Material Cost', 'Labor Cost'],
      ['Structure #1', 'Tangent', 46000, 22000],
      ['Structure #2', 'Small Running Angle', 65000, 87000],
      ['Structure #3', 'Medium Running Angle', 75000, 45000],
      ['Structure #4', 'Angle Deadend', 90000, 72000],
      ['Structure #5', 'Full Tension Deadend', 119000, 145000],
      [],
      ['345 kV - Steel Monopole'],
      ['Number of Circuits', '2 (default)'],
      ['ROW Width', '100 ft'],
      ['Structure Type', 'Name', 'Material Cost', 'Labor Cost'],
      ['Structure #1', 'Tangent', 98000, 60000],
      ['Structure #2', 'Small Running Angle', 135000, 100000],
      ['Structure #3', 'Medium Running Angle', 155000, 125000],
      ['Structure #4', 'Angle Deadend', 240000, 175000],
      ['Structure #5', 'Full Tension Deadend', 250000, 225000],
      [],
      ['345 kV - Steel Lattice'],
      ['Number of Circuits', '2 (default)'],
      ['ROW Width', '160 ft'],
      ['Structure Type', 'Name', 'Material Cost', 'Labor Cost'],
      ['Structure #1', 'Tangent', 70000, 110000],
      ['Structure #2', 'Small Running Angle', 101000, 162000],
      ['Structure #3', 'Medium Running Angle', 172000, 273000],
      ['Structure #4', 'Full Tension Deadend', 167000, 350000],
      ['Structure #5', 'N/A', 0, 0],
      [],
      ['765 kV - Steel Lattice'],
      ['Number of Circuits', '1 (default)'],
      ['ROW Width', '200 ft'],
      ['Structure Type', 'Name', 'Material Cost', 'Labor Cost'],
      ['Structure #1', 'Tangent', 154974, 227748],
      ['Structure #2', 'Small Running Angle', 218312, 329483],
      ['Structure #3', 'Angle Deadend', 428385, 632122],
      ['Structure #4', 'Full Tension Deadend', 527819, 1134719],
      ['Structure #5', 'N/A', 0, 0],
      [],
      ['Note: When 2 circuits are selected, Deadend structure costs are automatically doubled during calculations.']
    ];
    
    const wsDefaults = XLSX.utils.aoa_to_sheet(defaultsData);
    wsDefaults['!cols'] = [{ wch: 20 }, { wch: 30 }, { wch: 15 }, { wch: 15 }];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Structure Criteria');
    XLSX.utils.book_append_sheet(wb, wsDefaults, 'Defaults');
    XLSX.writeFile(wb, `Structure_Criteria_${voltage.replace(' ', '_')}_${structureFamily.replace(' ', '_')}.xlsx`);
  };

  const exportRatesTemplate = () => {
    const data = [
      ['Voltage', voltage],
      ['ROW Width (ft.)', rowWidth],
      [],
      ['Category', 'Sub-Category', 'Rate ($/mile)']
    ];
    
    // Get all $/mile variables
    const ratesData = [
      ['Right-of-Way and Land Acquisition', 'ROW Agent', rates.rowAgent || 0],
      ['Engineering and Design (Contract)', 'Land Surveying and SUE', rates.surveying || 0],
      ['Engineering and Design (Contract)', 'Aerial Survey', rates.aerialSurvey || 0],
      ['Engineering and Design (Contract)', 'Soil Borings', rates.soilBorings || 0],
      ['Engineering and Design (Contract)', 'Engineering and Design', rates.engineering || 0],
      ['Procurement of Material and Equipment', '1st Circuit Conductor Material', rates.conductor1Material || 0],
      ['Procurement of Material and Equipment', '2nd Circuit Conductor Material', rates.conductor2Material || 0],
      ['Procurement of Material and Equipment', '1st Circuit Shield Wire Material', rates.shieldWire1Material || 0],
      ['Procurement of Material and Equipment', '2nd Circuit Shield Wire Material', rates.shieldWire2Material || 0],
      ['Construction of Facilities (Contract)', 'Clearing', rates.clearing || 0],
      ['Construction of Facilities (Contract)', 'Environmental Controls', rates.environmental || 0],
      ['Construction of Facilities (Contract)', 'Damages', rates.damages || 0],
      ['Construction of Facilities (Contract)', 'Construction Management', rates.constructionMgmt || 0],
      ['Construction of Facilities (Contract)', 'Permitting', rates.permitting || 0],
      ['Construction of Facilities (Contract)', '1st Circuit Conductor Labor', rates.wireLabor || 0],
      ['Construction of Facilities (Contract)', '2nd Circuit Conductor Labor', rates.conductor2Labor || 0],
      ['Construction of Facilities (Contract)', '1st Circuit Shield Wire Labor', rates.shieldWire1Labor || 0],
      ['Construction of Facilities (Contract)', '2nd Circuit Shield Wire Labor', rates.shieldWire2Labor || 0]
    ];
    
    ratesData.forEach(row => data.push(row));
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 40 }, { wch: 35 }, { wch: 15 }];
    
    // Defaults sheet with all conductor/shield wire specifications
    const defaultsData = [
      ['Default Conductor and Shield Wire Specifications'],
      [],
      ['69 kV - 1 Circuit'],
      ['1st Circuit Conductor', '1-959 ACSS/TW "Suwannee"'],
      ['2nd Circuit Conductor', 'N/A'],
      ['1st Circuit Shield Wire', '69/138 kV OPGW'],
      ['2nd Circuit Shield Wire', 'N/A'],
      [],
      ['69 kV - 2 Circuits (Default)'],
      ['1st Circuit Conductor', '1-959 ACSS/TW "Suwannee"'],
      ['2nd Circuit Conductor', '1-959 ACSS/TW "Suwannee"'],
      ['1st Circuit Shield Wire', '69/138 kV OPGW'],
      ['2nd Circuit Shield Wire', '69/138 kV OPGW'],
      [],
      ['138 kV - 1 Circuit'],
      ['1st Circuit Conductor', '1-1926 ACSS/TW "Cumberland"'],
      ['2nd Circuit Conductor', 'N/A'],
      ['1st Circuit Shield Wire', '69/138 kV OPGW'],
      ['2nd Circuit Shield Wire', 'N/A'],
      [],
      ['138 kV - 2 Circuits (Default)'],
      ['1st Circuit Conductor', '1-1926 ACSS/TW "Cumberland"'],
      ['2nd Circuit Conductor', '1-1926 ACSS/TW "Cumberland"'],
      ['1st Circuit Shield Wire', '69/138 kV OPGW'],
      ['2nd Circuit Shield Wire', '69/138 kV OPGW'],
      [],
      ['345 kV - 1 Circuit'],
      ['1st Circuit Conductor', '2-1926 ACSS/TW "Cumberland"'],
      ['2nd Circuit Conductor', 'N/A'],
      ['1st Circuit Shield Wire', '345 kV OPGW'],
      ['2nd Circuit Shield Wire', 'N/A'],
      [],
      ['345 kV - 2 Circuits (Default)'],
      ['1st Circuit Conductor', '2-1926 ACSS/TW "Cumberland"'],
      ['2nd Circuit Conductor', '2-1926 ACSS/TW "Cumberland"'],
      ['1st Circuit Shield Wire', '345 kV OPGW'],
      ['2nd Circuit Shield Wire', '345 kV OPGW'],
      [],
      ['765 kV - 1 Circuit (Default)'],
      ['1st Circuit Conductor', '6-795 ACSR "Drake"'],
      ['2nd Circuit Conductor', 'N/A'],
      ['1st Circuit Shield Wire', '765 kV OPGW'],
      ['2nd Circuit Shield Wire', 'N/A'],
      [],
      ['ROW Width Defaults by Voltage'],
      [],
      ['Voltage', 'Structure Family', 'ROW Width (ft)', 'Default Circuits'],
      ['69 kV', 'Steel Monopole', '70', '2'],
      ['138 kV', 'Steel Monopole', '70', '2'],
      ['345 kV', 'Steel Monopole', '100', '2'],
      ['345 kV', 'Steel Lattice', '160', '2'],
      ['765 kV', 'Steel Lattice', '200', '1'],
      [],
      [],
      ['Default Rate Variables ($/mile)'],
      [],
      ['NOTE: Conductor and Shield Wire Material/Labor Rates are based on selected specifications'],
      [],
      ['1st Circuit Conductor Material Rates (by conductor type):'],
      ['1-795 ACSR "Drake"', '14,698'],
      ['1-959 ACSS/TW "Suwannee"', '17,711'],
      ['1-1926 ACSS/TW "Cumberland"', '34,952'],
      ['2-959 ACSS/TW "Suwannee"', '35,421'],
      ['2-1926 ACSS/TW "Cumberland"', '69,890'],
      ['6-795 ACSR "Drake"', '88,188'],
      [],
      ['1st Circuit Conductor Labor Rates (by conductor type):'],
      ['1-795 ACSR "Drake"', '15,258'],
      ['1-959 ACSS/TW "Suwannee"', '19,645'],
      ['1-1926 ACSS/TW "Cumberland"', '28,842'],
      ['2-959 ACSS/TW "Suwannee"', '28,985'],
      ['2-1926 ACSS/TW "Cumberland"', '54,111'],
      ['6-795 ACSR "Drake"', '87,018'],
      [],
      ['1st Circuit Shield Wire Material Rates (by shield wire type):'],
      ['3/8" EHS Steel', '1,958'],
      ['7/16" EHS Steel', '3,753'],
      ['69/138 kV OPGW', '6,635'],
      ['345 kV OPGW', '19,035'],
      ['765 kV OPGW', '19,035'],
      [],
      ['1st Circuit Shield Wire Labor Rates (by shield wire type):'],
      ['3/8" EHS Steel', '5,088'],
      ['7/16" EHS Steel', '5,088'],
      ['69/138 kV OPGW', '12,705'],
      ['345 kV OPGW', '13,977'],
      ['765 kV OPGW', '13,977'],
      [],
      ['Note: 2nd Circuit material/labor rates use the same rate structure as 1st Circuit based on selected type'],
      [],
      [],
      ['Other Rate Variables ($/mile)'],
      [],
      ['Category', 'Sub-Category', 'Default Rate'],
      [],
      [],
      ['Voltage-Dependent Rate Variables ($/mile)'],
      [],
      ['NOTE: All rates below are based on Voltage'],
      [],
      ['ROW Agent Rates (by Voltage):'],
      ['69 kV', '10,000'],
      ['138 kV', '10,000'],
      ['345 kV', '10,000'],
      ['765 kV', '30,000'],
      [],
      ['Land Surveying and SUE Rates (by Voltage):'],
      ['69 kV', '32,000'],
      ['138 kV', '32,000'],
      ['345 kV', '32,000'],
      ['765 kV', '75,000'],
      [],
      ['Aerial Survey Rates (by Voltage):'],
      ['69 kV', '4,500'],
      ['138 kV', '5,500'],
      ['345 kV', '5,500'],
      ['765 kV', '5,500'],
      [],
      ['Soil Borings Rates (by Voltage):'],
      ['69 kV', '30,000'],
      ['138 kV', '30,000'],
      ['345 kV', '15,000'],
      ['765 kV', '20,000'],
      [],
      ['Engineering and Design Rates (by Voltage):'],
      ['69 kV', '20,000'],
      ['138 kV', '20,000'],
      ['345 kV', '20,000'],
      ['765 kV', '100,000'],
      [],
      ['Environmental Controls Rates (by Voltage):'],
      ['69 kV', '150,000'],
      ['138 kV', '150,000'],
      ['345 kV', '150,000'],
      ['765 kV', '1,600,000'],
      [],
      ['Permitting Rates (by Voltage):'],
      ['69 kV', '5,000'],
      ['138 kV', '5,000'],
      ['345 kV', '10,000'],
      ['765 kV', '10,000'],
      [],
      ['Construction Management Rates (by Voltage):'],
      ['69 kV', '30,000'],
      ['138 kV', '30,000'],
      ['345 kV', '30,000'],
      ['765 kV', '60,000'],
      [],
      [],
      ['ROW Width-Dependent Rate Variables ($/mile)'],
      [],
      ['NOTE: Clearing and Damages rates are based on ROW Width'],
      [],
      ['Clearing Rates (by ROW Width):'],
      ['70 ft ROW', '15,000'],
      ['100 ft ROW', '15,000'],
      ['160 ft ROW', '20,000'],
      ['200 ft ROW', '20,000'],
      [],
      ['Damages Rates (by ROW Width):'],
      ['70 ft ROW', '6,500'],
      ['100 ft ROW', '6,500'],
      ['160 ft ROW', '10,000'],
      ['200 ft ROW', '30,000']
    ];
    
    const wsDefaults = XLSX.utils.aoa_to_sheet(defaultsData);
    wsDefaults['!cols'] = [{ wch: 40 }, { wch: 35 }, { wch: 15 }, { wch: 15 }];
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Rates Template');
    XLSX.utils.book_append_sheet(wb, wsDefaults, 'Defaults');
    XLSX.writeFile(wb, `Rates_Template_${voltage.replace(' ', '_')}_${rowWidth}ft.xlsx`);
  };

  const importRatesTemplate = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        // Read voltage and ROW width
        const importedVoltage = rows[0]?.[1] || '';
        const importedRowWidth = rows[1]?.[1]?.toString() || '';
        
        setVoltage(importedVoltage);
        setRowWidth(importedRowWidth);
        
        // Create mapping of subcategory names to rate keys
        const rateMapping = {
          'ROW Agent': 'rowAgent',
          'Land Surveying and SUE': 'surveying',
          'Aerial Survey': 'aerialSurvey',
          'Soil Borings': 'soilBorings',
          'Engineering and Design': 'engineering',
          '1st Circuit Conductor Material': 'conductor1Material',
          '2nd Circuit Conductor Material': 'conductor2Material',
          '1st Circuit Shield Wire Material': 'shieldWire1Material',
          '2nd Circuit Shield Wire Material': 'shieldWire2Material',
          'Clearing': 'clearing',
          'Environmental Controls': 'environmental',
          'Damages': 'damages',
          'Construction Management': 'constructionMgmt',
          'Permitting': 'permitting',
          '1st Circuit Conductor Labor': 'wireLabor',
          '2nd Circuit Conductor Labor': 'conductor2Labor',
          '1st Circuit Shield Wire Labor': 'shieldWire1Labor',
          '2nd Circuit Shield Wire Labor': 'shieldWire2Labor'
        };
        
        const newRates = { ...rates };
        
        // Read rates (starting from row 4, index 3)
        for (let i = 4; i < rows.length; i++) {
          const subCategoryName = rows[i]?.[1];
          const rateValue = parseFloat(rows[i]?.[2]) || 0;
          
          if (subCategoryName && rateMapping[subCategoryName]) {
            newRates[rateMapping[subCategoryName]] = rateValue;
          }
        }
        
        setRates(newRates);
        
        alert('Rates template imported successfully!');
      } catch (error) {
        alert('Error importing file. Please ensure it is a valid Rates Template Excel file.');
        console.error('Import error:', error);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const importStructureCriteria = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        // Read voltage and structure family
        const importedVoltage = rows[0]?.[1] || '';
        const importedFamily = rows[1]?.[1] || '';
        
        setVoltage(importedVoltage);
        setStructureFamily(importedFamily);
        
        // Read structure data (starting from row 4, index 3)
        const newNames = {};
        const newCosts = {};
        const newCustomMode = {};
        
        for (let i = 4; i < 9 && i < rows.length; i++) {
          const structNum = i - 3; // 1-5
          const name = rows[i]?.[1] || 'N/A';
          const material = parseFloat(rows[i]?.[2]) || 0;
          const labor = parseFloat(rows[i]?.[3]) || 0;
          
          newNames[`type${structNum}`] = name;
          newCosts[`type${structNum}Material`] = material;
          newCosts[`type${structNum}Labor`] = labor;
          newCustomMode[`type${structNum}`] = false;
        }
        
        setStructureNames(newNames);
        setStructureCosts(newCosts);
        setCustomMode(newCustomMode);
        
        alert('Structure criteria imported successfully!');
      } catch (error) {
        alert('Error importing file. Please ensure it is a valid Structure Criteria Excel file.');
        console.error('Import error:', error);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const exportRouteSummary = () => {
    // Build header row with route numbers
    const headers = ['Route ID'];
    const lengthRow = ['Route Length (mi)'];
    
    routes.forEach((route, idx) => {
      const routeLinks = links.filter(l => route.linkIds.includes(l.id));
      const totalLength = routeLinks.reduce((sum, l) => sum + (parseFloat(l.length) || 0), 0);
      
      // Extract route number from name (e.g., "Route 1" -> "1")
      const routeNumber = route.name.replace(/[^0-9]/g, '') || (idx + 1).toString();
      headers.push(routeNumber);
      lengthRow.push(totalLength.toFixed(2));
    });
    
    const data = [headers, lengthRow];
    
    // Calculate costs for all routes
    const allRouteCosts = routes.map(route => {
      const routeCategoryTotals = {};
      let transmissionSubtotal = 0;
      
      categories.slice(0, 7).forEach(mc => {
        if (subCategories[mc.key]) {
          let categoryTotal = 0;
          subCategories[mc.key].forEach((sc, si) => {
            let subCatTotal = 0;
            
            // Special handling for Land Acquisition - use route-based value
            if (mc.key === 'rightOfWay' && si === 0) {
              subCatTotal = parseFloat(routeLandAcquisition[route.id]) || 0;
            } else {
              route.linkIds.forEach(linkId => {
                const key = `${mc.key}_${si}_${linkId}`;
                subCatTotal += round(inputs[key]);
              });
            }
            
            categoryTotal += subCatTotal;
          });
          routeCategoryTotals[mc.key] = categoryTotal;
          transmissionSubtotal += categoryTotal;
        }
      });
      
      // Calculate escalation and contingency factors
      const totalYears = escalationYears + (escalationMonths / 12);
      const escalationFactor = Math.pow(1 + (escalationRate / 100), totalYears);
      const contingencyFactor = 1 + (contingency / 100);
      
      // Apply escalation and contingency to each category, then round up to nearest 1,000
      const routeCategoryTotalsAdjusted = {};
      let totalEscalation = 0;
      let totalContingency = 0;
      
      categories.slice(0, 7).forEach(mc => {
        const baseValue = routeCategoryTotals[mc.key] || 0;
        
        // Apply escalation
        const afterEscalation = baseValue * escalationFactor;
        const categoryEscalation = afterEscalation - baseValue;
        totalEscalation += categoryEscalation;
        
        // Apply contingency
        const afterContingency = afterEscalation * contingencyFactor;
        const categoryContingency = afterContingency - afterEscalation;
        totalContingency += categoryContingency;
        
        // Round final category value up to nearest 1,000
        routeCategoryTotalsAdjusted[mc.key] = Math.ceil(afterContingency / 1000) * 1000;
      });
      
      const routeEscalation = totalEscalation;
      const routeContingency = totalContingency;
      
      // Calculate transmission total from rounded category values
      const transmissionTotal = Object.values(routeCategoryTotalsAdjusted).reduce((sum, val) => sum + val, 0);
      const routeSubstationCost = Math.ceil((parseFloat(substationCost) || 0) / 1000) * 1000;
      const routeTotal = transmissionTotal + routeSubstationCost;
      
      return { routeCategoryTotalsAdjusted, transmissionSubtotal, routeEscalation, routeContingency, transmissionTotal, routeSubstationCost, routeTotal };
    });
    
    // Add category rows (adjusted values)
    categories.slice(0, 7).forEach(mc => {
      const row = [mc.label];
      allRouteCosts.forEach(routeCost => {
        row.push(routeCost.routeCategoryTotalsAdjusted[mc.key] || 0);
      });
      data.push(row);
    });
    
    // Add Transmission Line Cost
    const transmissionRow = ['Transmission Line Cost'];
    allRouteCosts.forEach(routeCost => {
      transmissionRow.push(routeCost.transmissionTotal);
    });
    data.push(transmissionRow);
    
    // Add Substation row
    const substationRow = ['Estimated Substation Facilities Cost'];
    allRouteCosts.forEach(routeCost => {
      substationRow.push(routeCost.routeSubstationCost);
    });
    data.push(substationRow);
    
    // Add Total Project Cost
    const totalRow = ['Total Project Cost'];
    allRouteCosts.forEach(routeCost => {
      totalRow.push(routeCost.routeTotal);
    });
    data.push(totalRow);
    
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{ wch: 50 }, ...routes.map(() => ({ wch: 15 }))];
    const wb = XLSX.utils.book_new();
    
    // Add Inputs Summary sheet first
    const inputsData = [
      ['INPUTS SUMMARY'],
      [],
      ['PROJECT SETUP'],
      ['Voltage', voltage],
      ['Structure Family', structureFamily],
      ['Number of Circuits', numCircuits],
      ['ROW Width (ft)', rowWidth],
      [],
      ['STRUCTURE TYPES'],
      ['Type 1', structureNames.type1, 'Material Cost', structureCosts.type1Material, 'Labor Cost', structureCosts.type1Labor],
      ['Type 2', structureNames.type2, 'Material Cost', structureCosts.type2Material, 'Labor Cost', structureCosts.type2Labor],
      ['Type 3', structureNames.type3, 'Material Cost', structureCosts.type3Material, 'Labor Cost', structureCosts.type3Labor],
      ['Type 4', structureNames.type4, 'Material Cost', structureCosts.type4Material, 'Labor Cost', structureCosts.type4Labor],
      ['Type 5', structureNames.type5, 'Material Cost', structureCosts.type5Material, 'Labor Cost', structureCosts.type5Labor],
      [],
      ['CONDUCTOR & SHIELD WIRE SPECIFICATIONS'],
      ['1st Circuit Conductor', conductor1Spec],
      ['2nd Circuit Conductor', conductor2Spec],
      ['1st Circuit Shield Wire', shieldWire1Spec],
      ['2nd Circuit Shield Wire', shieldWire2Spec],
      [],
      ['RATES ($/mile or %)'],
      ['ROW Agent', rates.rowAgent],
      ['Land Surveying and SUE', rates.surveying],
      ['Aerial Survey', rates.aerialSurvey],
      ['Soil Borings', rates.soilBorings],
      ['Engineering and Design', rates.engineering],
      ['Environmental Controls', rates.environmental],
      ['Clearing', rates.clearing],
      ['Damages', rates.damages],
      ['Permitting', rates.permitting],
      ['Construction Management', rates.constructionMgmt],
      ['1st Circuit Conductor Material', rates.conductor1Material],
      ['1st Circuit Conductor Labor', rates.wireLabor],
      ['2nd Circuit Conductor Material', rates.conductor2Material],
      ['2nd Circuit Conductor Labor', rates.conductor2Labor],
      ['1st Circuit Shield Wire Material', rates.shieldWire1Material],
      ['1st Circuit Shield Wire Labor', rates.shieldWire1Labor],
      ['2nd Circuit Shield Wire Material', rates.shieldWire2Material],
      ['2nd Circuit Shield Wire Labor', rates.shieldWire2Labor],
      ['M&S (Tax) %', rates.tax],
      ['P&S (Stores) %', rates.stores],
      ['Construction Overhead %', rates.overhead],
      [],
      ['ESCALATION & CONTINGENCY'],
      ['Escalation Rate (%)', escalationRate],
      ['Escalation Duration (Years)', escalationYears],
      ['Escalation Duration (Months)', escalationMonths],
      ['Contingency (%)', contingency],
      [],
      ['SUBSTATION'],
      ['Estimated Substation Facilities Cost', parseFloat(substationCost) || 0],
      [],
      ['LINKS'],
      ...links.map(link => {
        const structureInfo = [1, 2, 3, 4, 5]
          .filter(n => structureNames[`type${n}`] !== 'N/A')
          .map(n => `${structureNames[`type${n}`]}: ${link[`structureType${n}`] || 0}`)
          .join(', ');
        return [`Link ${link.link}`, `${link.length} mi`, structureInfo];
      })
    ];
    
    const wsInputs = XLSX.utils.aoa_to_sheet(inputsData);
    wsInputs['!cols'] = [{ wch: 35 }, { wch: 25 }, { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, wsInputs, 'Inputs Summary');
    
    // Add Route Link Combo sheet second
    const routeDetailsData = [['Route ID', 'Links', 'Length (mi)']];
    routes.forEach((route, idx) => {
      const routeLinks = links.filter(l => route.linkIds.includes(l.id));
      const totalLength = routeLinks.reduce((sum, l) => sum + (parseFloat(l.length) || 0), 0);
      const routeNumber = route.name.replace(/[^0-9]/g, '') || (idx + 1).toString();
      const linksStr = routeLinks.map(l => l.link).join('-');
      
      routeDetailsData.push([routeNumber, linksStr, totalLength.toFixed(2)]);
    });
    
    const wsDetails = XLSX.utils.aoa_to_sheet(routeDetailsData);
    wsDetails['!cols'] = [{ wch: 15 }, { wch: 30 }, { wch: 15 }];
    XLSX.utils.book_append_sheet(wb, wsDetails, 'Route Link Combo');
    
    // Add Route Costs Summary sheet third
    XLSX.utils.book_append_sheet(wb, ws, 'Route Costs Summary');
    
    XLSX.writeFile(wb, 'Route_Cost_Summary.xlsx');
  };

  const exportEquationsPDF = () => {
    // Create a complete HTML document with equations
    const equationsHTML = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Facility Calculator - Equations Reference</title>
  <style>
    @media print {
      body { margin: 0.5in; }
      .section { page-break-inside: avoid; }
    }
    body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; max-width: 8.5in; }
    h1 { color: #1e40af; border-bottom: 3px solid #1e40af; padding-bottom: 10px; }
    h2 { color: #2563eb; margin-top: 30px; border-bottom: 2px solid #3b82f6; padding-bottom: 5px; }
    h3 { color: #3b82f6; margin-top: 20px; }
    .equation { background: #f1f5f9; padding: 15px; margin: 10px 0; border-left: 4px solid #3b82f6; font-family: 'Courier New', monospace; }
    .note { background: #fef3c7; padding: 10px; margin: 10px 0; border-left: 4px solid #f59e0b; }
    .section { margin-bottom: 30px; }
    ul { margin-left: 20px; }
    li { margin: 5px 0; }
    .print-instructions { background: #e0f2fe; padding: 15px; margin: 20px 0; border: 2px solid #0284c7; border-radius: 5px; }
  </style>
</head>
<body>
  <div class="print-instructions">
    <strong>📄 To save as PDF:</strong> Press Ctrl+P (Windows) or Cmd+P (Mac), then select "Save as PDF" as the destination.
  </div>

  <h1>Transmission Line Facility Cost Calculator - Equations Reference</h1>
  <p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
  
  <div class="section">
    <h2>1. Structure Costs</h2>
    <h3>1.1 Structure Material Cost per Link</h3>
    <div class="equation">
      Structure Material Cost = Σ(Structure Type Quantity × Structure Type Material Cost)
    </div>
    <p>Where the sum includes all structure types (1-5) for each link.</p>
    <div class="note">
      <strong>Note:</strong> For 2-circuit configurations, Angle Deadend and Full Tension Deadend costs are doubled.
    </div>

    <h3>1.2 Structure Labor Cost per Link</h3>
    <div class="equation">
      Structure Labor Cost = Σ(Structure Type Quantity × Structure Type Labor Cost)
    </div>
    <p>Where the sum includes all structure types (1-5) for each link.</p>
    <div class="note">
      <strong>Note:</strong> For 2-circuit configurations, Angle Deadend and Full Tension Deadend costs are doubled.
    </div>
  </div>

  <div class="section">
    <h2>2. Rate-Based Costs</h2>
    <h3>2.1 Fixed Cost Sub-Categories</h3>
    <div class="equation">
      Fixed Cost = User-Defined Fixed Amount
    </div>
    <p>Examples: Land Acquisition (route-based fixed value)</p>

    <h3>2.2 Per-Mile Cost Sub-Categories</h3>
    <div class="equation">
      Per-Mile Cost = Link Length (miles) × Rate ($/mile)
    </div>
    <p>Examples: ROW Agent, Land Surveying, Aerial Survey, Engineering, Environmental Controls, Clearing, Damages, Permitting, Construction Management, Conductor Materials, Conductor Labor, Shield Wire Materials, Shield Wire Labor</p>

    <h3>2.3 Percentage-Based Cost Sub-Categories</h3>
    <div class="equation">
      Percentage Cost = Base Cost × (Rate / 100)
    </div>
    <p>Examples:</p>
    <ul>
      <li><strong>M&S (Tax):</strong> Applied to all material costs (Structure Material, Conductor Material, Shield Wire Material)</li>
      <li><strong>P&S (Stores):</strong> Applied to all material costs (Structure Material, Conductor Material, Shield Wire Material)</li>
      <li><strong>Construction Overhead:</strong> Applied to sum of categories 0-6 including Permitting</li>
    </ul>
  </div>

  <div class="section">
    <h2>3. Category Totals</h2>
    <h3>3.1 Link-Level Category Total</h3>
    <div class="equation">
      Category Total = Σ(Sub-Category Costs within Category)
    </div>
    <p>Seven main categories:</p>
    <ol>
      <li>Right-of-Way (Utility)</li>
      <li>Procurement of Material (Utility)</li>
      <li>Engineering and Design (Contract)</li>
      <li>Environmental Controls (Contract)</li>
      <li>Construction of Facilities (Contract)</li>
      <li>Construction Management (Contract)</li>
      <li>Construction of Substation (Utility)</li>
    </ol>

    <h3>3.2 Route-Level Category Total</h3>
    <div class="equation">
      Route Category Total = Σ(Link Category Totals for all links in route)
    </div>
    <div class="note">
      <strong>Exception:</strong> Land Acquisition uses route-based fixed value instead of sum of link values.
    </div>
  </div>

  <div class="section">
    <h2>4. Escalation</h2>
    <h3>4.1 Escalation Factor</h3>
    <div class="equation">
      Total Years = Escalation Years + (Escalation Months / 12)
      <br><br>
      Escalation Factor = (1 + Escalation Rate / 100) ^ Total Years
    </div>

    <h3>4.2 Apply Escalation to Category</h3>
    <div class="equation">
      Escalated Category Cost = Base Category Cost × Escalation Factor
    </div>
  </div>

  <div class="section">
    <h2>5. Contingency</h2>
    <h3>5.1 Contingency Factor</h3>
    <div class="equation">
      Contingency Factor = 1 + (Contingency Rate / 100)
    </div>

    <h3>5.2 Apply Contingency to Escalated Category</h3>
    <div class="equation">
      Final Category Cost = Escalated Category Cost × Contingency Factor
    </div>
  </div>

  <div class="section">
    <h2>6. Rounding</h2>
    <h3>6.1 Category-Level Rounding</h3>
    <div class="equation">
      Rounded Category Cost = CEILING(Final Category Cost / 1000) × 1000
    </div>
    <p>Each category is rounded UP to the nearest $1,000 after escalation and contingency are applied.</p>
  </div>

  <div class="section">
    <h2>7. Transmission Line Cost</h2>
    <h3>7.1 Link Transmission Line Cost</h3>
    <div class="equation">
      Link Transmission Cost = Σ(Rounded Category Costs for categories 0-6)
    </div>

    <h3>7.2 Route Transmission Line Cost</h3>
    <div class="equation">
      Route Transmission Cost = Σ(Rounded Category Costs for categories 0-6)
    </div>
    <p>Where each category is the sum of that category across all links in the route, then escalated, contingency applied, and rounded.</p>
  </div>

  <div class="section">
    <h2>8. Total Project Cost</h2>
    <h3>8.1 Route Total Project Cost</h3>
    <div class="equation">
      Route Total = Route Transmission Cost + Substation Cost
    </div>
    <p>Where:</p>
    <ul>
      <li><strong>Route Transmission Cost:</strong> Sum of 7 rounded category costs (categories 0-6)</li>
      <li><strong>Substation Cost:</strong> Estimated Substation Facilities Cost (rounded to nearest $1,000)</li>
    </ul>
  </div>

  <div class="section">
    <h2>9. Cost per Mile</h2>
    <h3>9.1 Route Average Cost per Mile</h3>
    <div class="equation">
      Avg. $/mi = Route Transmission Cost / Route Total Length
    </div>
    <p>Where Route Total Length is the sum of all link lengths in the route.</p>
  </div>

  <div class="section">
    <h2>10. Calculation Order</h2>
    <p><strong>For each link or route:</strong></p>
    <ol>
      <li>Calculate base sub-category costs (structure costs, rate-based costs)</li>
      <li>Sum sub-categories into category totals</li>
      <li>Apply escalation to each category</li>
      <li>Apply contingency to each escalated category</li>
      <li>Round each category up to nearest $1,000</li>
      <li>Sum all rounded categories to get Transmission Line Cost</li>
      <li>Add Substation Cost to get Total Project Cost</li>
    </ol>
  </div>

  <div class="section">
    <h2>11. Default Rates</h2>
    <p>The calculator auto-populates default rates based on:</p>
    <ul>
      <li><strong>Voltage-Dependent:</strong> ROW Agent, Land Surveying & SUE, Aerial Survey, Soil Borings, Engineering & Design, Environmental Controls, Permitting, Construction Management</li>
      <li><strong>ROW Width-Dependent:</strong> Clearing, Damages</li>
      <li><strong>Conductor Specification-Dependent:</strong> Conductor Material and Labor rates (both circuits)</li>
      <li><strong>Shield Wire Specification-Dependent:</strong> Shield Wire Material and Labor rates (both circuits)</li>
    </ul>
    <p>All default rates can be manually overridden in Step 3 (Link Costs).</p>
  </div>

  <p style="margin-top: 50px; text-align: center; color: #64748b; font-size: 12px;">
    Transmission Line Facility Cost Calculator © ${new Date().getFullYear()}
  </p>

  <script>
    // Auto-trigger print dialog when page loads
    window.onload = function() {
      window.print();
    };
  </script>
</body>
</html>
    `;

    // Create blob and download link
    const blob = new Blob([equationsHTML], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Equations_Reference.html';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    // Also open in new window for immediate printing
    setTimeout(() => {
      const printWindow = window.open('', '_blank');
      if (printWindow) {
        printWindow.document.write(equationsHTML);
        printWindow.document.close();
      }
    }, 100);
  };

  const exportStatistics = () => {
    // Calculate route statistics
    const routeStats = routes.map(route => {
      const routeLinks = links.filter(l => route.linkIds.includes(l.id));
      const totalLength = routeLinks.reduce((sum, l) => sum + (parseFloat(l.length) || 0), 0);
      
      // Calculate transmission line cost for this route
      const routeCategoryTotals = {};
      categories.slice(0, 7).forEach(mc => {
        if (subCategories[mc.key]) {
          let categoryTotal = 0;
          subCategories[mc.key].forEach((sc, si) => {
            let subCatTotal = 0;
            if (mc.key === 'rightOfWay' && si === 0) {
              subCatTotal = parseFloat(routeLandAcquisition[route.id]) || 0;
            } else {
              route.linkIds.forEach(linkId => {
                const key = `${mc.key}_${si}_${linkId}`;
                subCatTotal += round(inputs[key]);
              });
            }
            categoryTotal += subCatTotal;
          });
          routeCategoryTotals[mc.key] = categoryTotal;
        }
      });
      
      const totalYears = escalationYears + (escalationMonths / 12);
      const escalationFactor = Math.pow(1 + (escalationRate / 100), totalYears);
      const contingencyFactor = 1 + (contingency / 100);
      
      const routeCategoryTotalsAdjusted = {};
      categories.slice(0, 7).forEach(mc => {
        const baseValue = routeCategoryTotals[mc.key] || 0;
        const afterEscalation = baseValue * escalationFactor;
        const afterContingency = afterEscalation * contingencyFactor;
        routeCategoryTotalsAdjusted[mc.key] = Math.ceil(afterContingency / 1000) * 1000;
      });
      
      const transmissionTotal = Object.values(routeCategoryTotalsAdjusted).reduce((sum, val) => sum + val, 0);
      
      return {
        name: route.name,
        length: totalLength,
        cost: transmissionTotal
      };
    });

    // Calculate statistics for bell curve
    const costs = routeStats.map(r => r.cost);
    const mean = costs.reduce((sum, c) => sum + c, 0) / costs.length;
    const variance = costs.reduce((sum, c) => sum + Math.pow(c - mean, 2), 0) / costs.length;
    const stdDev = Math.sqrt(variance);
    
    // Create HTML with charts
    const statsHTML = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Route Statistics - Charts</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <style>
    @media print {
      body { margin: 0.5in; }
      .chart-container { page-break-inside: avoid; }
    }
    body { 
      font-family: Arial, sans-serif; 
      margin: 40px; 
      background: #f8fafc;
    }
    h1 { 
      color: #1e40af; 
      border-bottom: 3px solid #1e40af; 
      padding-bottom: 10px; 
    }
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 20px;
      margin: 30px 0;
    }
    .stat-card {
      background: white;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      border-left: 4px solid #3b82f6;
    }
    .stat-label {
      font-size: 14px;
      color: #64748b;
      margin-bottom: 5px;
    }
    .stat-value {
      font-size: 24px;
      font-weight: bold;
      color: #1e293b;
    }
    .chart-container {
      background: white;
      padding: 30px;
      margin: 30px 0;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .chart-title {
      font-size: 18px;
      font-weight: bold;
      color: #1e293b;
      margin-bottom: 20px;
    }
    canvas {
      max-height: 400px;
    }
    .print-instructions {
      background: #e0f2fe;
      padding: 15px;
      margin: 20px 0;
      border: 2px solid #0284c7;
      border-radius: 5px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
      background: white;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    th, td {
      padding: 12px;
      text-align: left;
      border-bottom: 1px solid #e2e8f0;
    }
    th {
      background: #f1f5f9;
      font-weight: bold;
      color: #1e293b;
    }
  </style>
</head>
<body>
  <div class="print-instructions">
    <strong>📊 To save as PDF:</strong> Press Ctrl+P (Windows) or Cmd+P (Mac), then select "Save as PDF" as the destination.
  </div>

  <h1>Route Statistics Analysis</h1>
  <p><strong>Date:</strong> ${new Date().toLocaleDateString()}</p>
  
  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-label">Total Routes</div>
      <div class="stat-value">${routeStats.length}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Average Cost</div>
      <div class="stat-value">$${Math.round(mean).toLocaleString('en-US')}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Std Deviation</div>
      <div class="stat-value">$${Math.round(stdDev).toLocaleString('en-US')}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Min Cost</div>
      <div class="stat-value">$${Math.min(...costs).toLocaleString('en-US')}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Max Cost</div>
      <div class="stat-value">$${Math.max(...costs).toLocaleString('en-US')}</div>
    </div>
  </div>

  <h2>Route Data Summary</h2>
  <table>
    <thead>
      <tr>
        <th>Route</th>
        <th>Length (mi)</th>
        <th>Transmission Line Cost</th>
        <th>Cost per Mile</th>
      </tr>
    </thead>
    <tbody>
      ${routeStats.map(r => `
        <tr>
          <td>Route ${r.name}</td>
          <td>${r.length.toFixed(2)}</td>
          <td>$${r.cost.toLocaleString('en-US')}</td>
          <td>$${r.length > 0 ? Math.round(r.cost / r.length).toLocaleString('en-US') : '0'}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>

  <div class="chart-container">
    <div class="chart-title">Transmission Line Cost Distribution (Normal Distribution Curve)</div>
    <canvas id="bellCurveChart"></canvas>
  </div>

  <div class="chart-container">
    <div class="chart-title">Cost vs. Route Length (Scatter Plot)</div>
    <canvas id="scatterChart"></canvas>
  </div>

  <script>
    // Bell Curve Chart
    const costs = ${JSON.stringify(costs)};
    const mean = ${mean};
    const stdDev = ${stdDev};
    
    // Generate normal distribution curve
    const min = Math.min(...costs) - stdDev;
    const max = Math.max(...costs) + stdDev;
    const step = (max - min) / 100;
    const curveData = [];
    
    for (let x = min; x <= max; x += step) {
      const y = (1 / (stdDev * Math.sqrt(2 * Math.PI))) * 
                Math.exp(-0.5 * Math.pow((x - mean) / stdDev, 2));
      curveData.push({ x: x, y: y });
    }
    
    // Create histogram bins
    const binCount = Math.min(10, costs.length);
    const binSize = (max - min) / binCount;
    const bins = Array(binCount).fill(0);
    const binLabels = [];
    
    costs.forEach(cost => {
      const binIndex = Math.min(Math.floor((cost - min) / binSize), binCount - 1);
      bins[binIndex]++;
    });
    
    for (let i = 0; i < binCount; i++) {
      const binStart = min + (i * binSize);
      binLabels.push('$' + Math.round(binStart / 1000) + 'K');
    }
    
    // Normalize bins to match curve scale
    const maxBin = Math.max(...bins);
    const maxCurve = Math.max(...curveData.map(d => d.y));
    const normalizedBins = bins.map(b => (b / maxBin) * maxCurve);
    
    new Chart(document.getElementById('bellCurveChart'), {
      type: 'line',
      data: {
        labels: binLabels,
        datasets: [
          {
            type: 'bar',
            label: 'Actual Route Distribution',
            data: normalizedBins,
            backgroundColor: 'rgba(59, 130, 246, 0.5)',
            borderColor: 'rgba(59, 130, 246, 1)',
            borderWidth: 1
          },
          {
            label: 'Normal Distribution Curve',
            data: curveData.map(d => d.y),
            borderColor: 'rgba(239, 68, 68, 1)',
            backgroundColor: 'rgba(239, 68, 68, 0.1)',
            borderWidth: 2,
            fill: true,
            tension: 0.4,
            pointRadius: 0
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: {
            display: true,
            position: 'top'
          },
          title: {
            display: true,
            text: 'Mean: $' + Math.round(mean).toLocaleString() + ' | Std Dev: $' + Math.round(stdDev).toLocaleString()
          }
        },
        scales: {
          x: {
            title: {
              display: true,
              text: 'Transmission Line Cost'
            }
          },
          y: {
            title: {
              display: true,
              text: 'Frequency (Normalized)'
            }
          }
        }
      }
    });

    // Scatter Plot Chart
    const scatterData = ${JSON.stringify(routeStats.map(r => ({ x: r.length, y: r.cost, label: 'Route ' + r.name })))};
    
    new Chart(document.getElementById('scatterChart'), {
      type: 'scatter',
      data: {
        datasets: [{
          label: 'Routes',
          data: scatterData,
          backgroundColor: 'rgba(59, 130, 246, 0.6)',
          borderColor: 'rgba(59, 130, 246, 1)',
          borderWidth: 2,
          pointRadius: 8,
          pointHoverRadius: 12
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: {
            display: true
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                return context.raw.label + ': ' + 
                       context.parsed.x.toFixed(2) + ' mi, $' + 
                       context.parsed.y.toLocaleString();
              }
            }
          }
        },
        scales: {
          x: {
            title: {
              display: true,
              text: 'Route Length (miles)'
            },
            beginAtZero: true
          },
          y: {
            title: {
              display: true,
              text: 'Transmission Line Cost ($)'
            },
            beginAtZero: true,
            ticks: {
              callback: function(value) {
                return '$' + (value / 1000000).toFixed(1) + 'M';
              }
            }
          }
        }
      }
    });

    // Auto-print when loaded
    window.onload = function() {
      setTimeout(() => window.print(), 500);
    };
  </script>
</body>
</html>
    `;

    // Create blob and download
    const blob = new Blob([statsHTML], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Route_Statistics.html';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    
    // Open in new window
    setTimeout(() => {
      const printWindow = window.open('', '_blank');
      if (printWindow) {
        printWindow.document.write(statsHTML);
        printWindow.document.close();
      }
    }, 100);
  };

  const reset = () => {
    const confirmed = window.confirm('Are you sure you want to start over? All entry data will be cleared.');
    if (!confirmed) {
      return;
    }
    
    const fixed = {};
    const methods = {};
    categories.forEach(c => { 
      if (c.fixed) {
        fixed[c.key] = [...c.fixed];
        c.fixed.forEach((subcat, idx) => {
          const key = `${c.key}_${idx}`;
          if (c.key === 'rightOfWay' && idx === 1) methods[key] = '$/mile';
          else if (c.key === 'engineeringContract') methods[key] = '$/mile';
          else if (c.key === 'procurement' && idx >= 1 && idx <= 4) methods[key] = '$/mile'; // All conductor and shield wire materials
          else if (c.key === 'procurement' && (idx === 5 || idx === 6)) methods[key] = '%';
          else if (c.key === 'transmissionContract' && idx >= 0 && idx <= 4) methods[key] = '$/mile'; // Clearing through Permitting
          else if (c.key === 'transmissionContract' && idx === 5) methods[key] = 'fixed'; // Structure Labor
          else if (c.key === 'transmissionContract' && idx >= 6 && idx <= 9) methods[key] = '$/mile'; // All conductor/shield wire labor
          else if (c.key === 'transmissionContract' && idx === 10) methods[key] = '%';
          else methods[key] = 'fixed';
        });
      }
    });
    setStep(0);
    setVoltage('');
    setRowWidth('');
    setNumCircuits('1');
    setConductor1Spec('');
    setConductor2Spec('');
    setShieldWire1Spec('');
    setShieldWire2Spec('');
    setStructureFamily('');
    setStructureNames({
      type1: 'Tangent',
      type2: 'Medium Running Angle',
      type3: 'Angle Deadend',
      type4: 'Full Tension Deadend',
      type5: 'N/A'
    });
    setCustomMode({
      type1: false,
      type2: false,
      type3: false,
      type4: false,
      type5: false
    });
    setStructureCosts({
      type1Material: 0,
      type1Labor: 0,
      type2Material: 0,
      type2Labor: 0,
      type3Material: 0,
      type3Labor: 0,
      type4Material: 0,
      type4Labor: 0,
      type5Material: 0,
      type5Labor: 0
    });
    setSubCategories(fixed);
    setSubCategoryMethods(methods);
    setInputs({});
    setTotal(null);
    setRoutes([]);
    setSubstationCost('');
    setContingency(15);
    setEscalationRate(3);
    setEscalationYears(2);
    setEscalationMonths(0);
    setEditingCat(null);
    setTempSubCat('');
    setRates({ tax: 8.25, stores: 5.0, overhead: 8.275, rowAgent: 15000, surveying: 0, aerialSurvey: 0, soilBorings: 0, engineering: 0, clearing: 0, environmental: 0, damages: 0, constructionMgmt: 0, permitting: 0, wireLabor: 0, conductor2Labor: 0, shieldWire1Labor: 0, shieldWire2Labor: 0, conductor1Material: 0, conductor2Material: 0, shieldWire1Material: 0, shieldWire2Material: 0 });
  };

  const Input = ({ type = "text", className = "", ...props }) => <input type={type} className={`w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 ${className}`} style={{WebkitAppearance: 'none', MozAppearance: 'none'}} {...props} />;
  const Btn = ({ children, ...props }) => <button className="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-6 rounded-lg" {...props}>{children}</button>;

  if (step === 0) return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
      <div className="w-full max-w-4xl bg-slate-800 rounded-2xl p-8 border border-slate-700">
        <h1 className="text-3xl font-bold text-white mb-6">Select Structure Types</h1>
        <div className="mb-6 flex justify-between items-center">
          <div className="flex gap-3">
            <button 
              onClick={exportStructureCriteria}
              className="bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-4 rounded-lg"
            >
              Export Criteria
            </button>
            <label className="bg-purple-600 hover:bg-purple-700 text-white font-semibold py-2 px-4 rounded-lg cursor-pointer">
              Import Criteria
              <input 
                type="file" 
                accept=".xlsx,.xls"
                className="hidden"
                onChange={e => {
                  if (e.target.files?.[0]) {
                    importStructureCriteria(e.target.files[0]);
                    e.target.value = ''; // Reset input
                  }
                }}
              />
            </label>
          </div>
          <Btn onClick={() => setStep(1)}>Continue</Btn>
        </div>
        <p className="text-slate-300 mb-6">Define the voltage, structure family, and costs for your structure types.</p>
        
        {/* Voltage Selection */}
        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <label className="block text-slate-300 text-sm mb-2 font-semibold">Voltage</label>
          <select
            className="w-full max-w-md px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
            value={voltage}
            onChange={e => {
              setVoltage(e.target.value);
              
              // Set default ROW Agent rate based on voltage
              const rowAgentRates = {
                '69 kV': 10000,
                '138 kV': 10000,
                '345 kV': 10000,
                '765 kV': 30000
              };
              // Set default Construction Management rate based on voltage
              const constructionMgmtRates = {
                '69 kV': 30000,
                '138 kV': 30000,
                '345 kV': 30000,
                '765 kV': 60000
              };
              // Set default Engineering and Design rate based on voltage
              const engineeringRates = {
                '69 kV': 20000,
                '138 kV': 20000,
                '345 kV': 20000,
                '765 kV': 100000
              };
              // Set default Environmental Controls rate based on voltage
              const environmentalRates = {
                '69 kV': 150000,
                '138 kV': 150000,
                '345 kV': 150000,
                '765 kV': 1600000
              };
              // Set default Aerial Survey rate based on voltage
              const aerialSurveyRates = {
                '69 kV': 4500,
                '138 kV': 5500,
                '345 kV': 5500,
                '765 kV': 5500
              };
              // Set default Permitting rate based on voltage
              const permittingRates = {
                '69 kV': 5000,
                '138 kV': 5000,
                '345 kV': 10000,
                '765 kV': 10000
              };
              // Set default Land Surveying and SUE rate based on voltage
              const surveyingRates = {
                '69 kV': 32000,
                '138 kV': 32000,
                '345 kV': 32000,
                '765 kV': 75000
              };
              // Set default Soil Borings rate based on voltage
              const soilBoringsRates = {
                '69 kV': 30000,
                '138 kV': 30000,
                '345 kV': 15000,
                '765 kV': 20000
              };
              if (rowAgentRates[e.target.value]) {
                setRates({
                  ...rates, 
                  rowAgent: rowAgentRates[e.target.value],
                  constructionMgmt: constructionMgmtRates[e.target.value],
                  engineering: engineeringRates[e.target.value],
                  environmental: environmentalRates[e.target.value],
                  aerialSurvey: aerialSurveyRates[e.target.value],
                  permitting: permittingRates[e.target.value],
                  surveying: surveyingRates[e.target.value],
                  soilBorings: soilBoringsRates[e.target.value]
                });
              }
              
              // Auto-set number of circuits based on voltage
              if (e.target.value === '765 kV') {
                setNumCircuits('1');
                // Set 2nd circuit fields to N/A for 1 circuit
                setConductor2Spec('N/A');
                setShieldWire2Spec('N/A');
                // Auto-populate conductor and shield wire specs for 765 kV - 1 circuit
                setConductor1Spec('6-795 ACSR "Drake"');
                setShieldWire1Spec('765 kV OPGW');
                // Set conductor and shield wire rates for 765 kV
                setRates(prev => ({
                  ...prev,
                  conductor1Material: 88188,
                  wireLabor: 87018,
                  shieldWire1Material: 19035,
                  shieldWire1Labor: 13977
                }));
              } else if (e.target.value === '69 kV' || e.target.value === '138 kV' || e.target.value === '345 kV') {
                setNumCircuits('2');
                // Clear N/A when setting to 2 circuits
                if (conductor2Spec === 'N/A') setConductor2Spec('');
                if (shieldWire2Spec === 'N/A') setShieldWire2Spec('');
                
                // Auto-populate conductor and shield wire specs for 69 kV - 2 circuits
                if (e.target.value === '69 kV') {
                  setConductor1Spec('1-959 ACSS/TW "Suwannee"');
                  setConductor2Spec('1-959 ACSS/TW "Suwannee"');
                  setShieldWire1Spec('69/138 kV OPGW');
                  setShieldWire2Spec('69/138 kV OPGW');
                  // Set conductor and shield wire rates for 69 kV
                  setRates(prev => ({
                    ...prev,
                    conductor1Material: 17711,
                    wireLabor: 19645,
                    conductor2Material: 17711,
                    conductor2Labor: 19645,
                    shieldWire1Material: 6635,
                    shieldWire1Labor: 12705,
                    shieldWire2Material: 6635,
                    shieldWire2Labor: 12705
                  }));
                }
                // Auto-populate conductor and shield wire specs for 138 kV - 2 circuits
                else if (e.target.value === '138 kV') {
                  setConductor1Spec('1-1926 ACSS/TW "Cumberland"');
                  setConductor2Spec('1-1926 ACSS/TW "Cumberland"');
                  setShieldWire1Spec('69/138 kV OPGW');
                  setShieldWire2Spec('69/138 kV OPGW');
                  // Set conductor and shield wire rates for 138 kV
                  setRates(prev => ({
                    ...prev,
                    conductor1Material: 34952,
                    wireLabor: 28842,
                    conductor2Material: 34952,
                    conductor2Labor: 28842,
                    shieldWire1Material: 6635,
                    shieldWire1Labor: 12705,
                    shieldWire2Material: 6635,
                    shieldWire2Labor: 12705
                  }));
                }
                // Auto-populate conductor and shield wire specs for 345 kV - 2 circuits
                else if (e.target.value === '345 kV') {
                  setConductor1Spec('2-1926 ACSS/TW "Cumberland"');
                  setConductor2Spec('2-1926 ACSS/TW "Cumberland"');
                  setShieldWire1Spec('345 kV OPGW');
                  setShieldWire2Spec('345 kV OPGW');
                  // Set conductor and shield wire rates for 345 kV
                  setRates(prev => ({
                    ...prev,
                    conductor1Material: 69890,
                    wireLabor: 54111,
                    conductor2Material: 69890,
                    conductor2Labor: 54111,
                    shieldWire1Material: 19035,
                    shieldWire1Labor: 13977,
                    shieldWire2Material: 19035,
                    shieldWire2Labor: 13977
                  }));
                }
              }
              
              if (e.target.value === '765 kV') {
                setStructureFamily('Steel Lattice');
                setRowWidth('200');
                // Set Damages and Clearing rates for 200 ft ROW
                setRates(prev => ({...prev, damages: 30000, clearing: 20000}));
                // Auto-populate structure types and costs for 765 kV Steel Lattice
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Angle Deadend',
                  type4: 'Full Tension Deadend',
                  type5: 'N/A'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 154974,
                  type1Labor: 227748,
                  type2Material: 218312,
                  type2Labor: 329483,
                  type3Material: 428385,
                  type3Labor: 632122,
                  type4Material: 527819,
                  type4Labor: 1134719,
                  type5Material: 0,
                  type5Labor: 0
                });
              } else if (e.target.value === '69 kV' || e.target.value === '138 kV') {
                setStructureFamily('Steel Monopole');
                setRowWidth('70');
                // Set Damages and Clearing rates for 70 ft ROW
                setRates(prev => ({...prev, damages: 6500, clearing: 15000}));
              }
              
              // Auto-populate structure types and costs for 69 kV or 138 kV Steel Monopole
              if ((e.target.value === '69 kV' || e.target.value === '138 kV') && (structureFamily === 'Steel Monopole' || e.target.value === '69 kV' || e.target.value === '138 kV')) {
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Medium Running Angle',
                  type4: 'Angle Deadend',
                  type5: 'Full Tension Deadend'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 46000,
                  type1Labor: 22000,
                  type2Material: 65000,
                  type2Labor: 87000,
                  type3Material: 75000,
                  type3Labor: 45000,
                  type4Material: 90000,
                  type4Labor: 72000,
                  type5Material: 119000,
                  type5Labor: 145000
                });
              }
              // Auto-populate structure types for 345 kV Steel Monopole (no default costs yet)
              else if (e.target.value === '345 kV' && (structureFamily === 'Steel Monopole' || e.target.value === '345 kV')) {
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Medium Running Angle',
                  type4: 'Angle Deadend',
                  type5: 'Full Tension Deadend'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 98000,
                  type1Labor: 60000,
                  type2Material: 135000,
                  type2Labor: 100000,
                  type3Material: 155000,
                  type3Labor: 125000,
                  type4Material: 240000,
                  type4Labor: 175000,
                  type5Material: 250000,
                  type5Labor: 225000
                });
              }
              // Auto-populate structure types for 345 kV Steel Lattice
              else if (e.target.value === '345 kV' && structureFamily === 'Steel Lattice') {
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Medium Running Angle',
                  type4: 'Full Tension Deadend',
                  type5: 'N/A'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 70000,
                  type1Labor: 110000,
                  type2Material: 101000,
                  type2Labor: 162000,
                  type3Material: 172000,
                  type3Labor: 273000,
                  type4Material: 167000,
                  type4Labor: 350000,
                  type5Material: 0,
                  type5Labor: 0
                });
              }
            }}
          >
            <option value="">Select Voltage</option>
            <option value="69 kV">69 kV</option>
            <option value="138 kV">138 kV</option>
            <option value="345 kV">345 kV</option>
            <option value="765 kV">765 kV</option>
          </select>
        </div>

        {/* Structure Family Selection */}
        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <label className="block text-slate-300 text-sm mb-2 font-semibold">Structure Family</label>
          <select
            className="w-full max-w-md px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
            value={structureFamily}
            onChange={e => {
              setStructureFamily(e.target.value);
              // Auto-populate structure types and costs for 69 kV or 138 kV Steel Monopole
              if ((voltage === '69 kV' || voltage === '138 kV') && e.target.value === 'Steel Monopole') {
                setRowWidth('70');
                // Set Damages and Clearing rates for 70 ft ROW
                setRates(prev => ({...prev, damages: 6500, clearing: 15000}));
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Medium Running Angle',
                  type4: 'Angle Deadend',
                  type5: 'Full Tension Deadend'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 46000,
                  type1Labor: 22000,
                  type2Material: 65000,
                  type2Labor: 87000,
                  type3Material: 75000,
                  type3Labor: 45000,
                  type4Material: 90000,
                  type4Labor: 72000,
                  type5Material: 119000,
                  type5Labor: 145000
                });
              }
              // Auto-populate structure types for 345 kV Steel Monopole (no default costs yet)
              else if (voltage === '345 kV' && e.target.value === 'Steel Monopole') {
                setRowWidth('100');
                // Set Damages and Clearing rates for 100 ft ROW
                setRates(prev => ({...prev, damages: 6500, clearing: 15000}));
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Medium Running Angle',
                  type4: 'Angle Deadend',
                  type5: 'Full Tension Deadend'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 98000,
                  type1Labor: 60000,
                  type2Material: 135000,
                  type2Labor: 100000,
                  type3Material: 155000,
                  type3Labor: 125000,
                  type4Material: 240000,
                  type4Labor: 175000,
                  type5Material: 250000,
                  type5Labor: 225000
                });
              }
              // Auto-populate structure types and costs for 345 kV Steel Lattice
              else if (voltage === '345 kV' && e.target.value === 'Steel Lattice') {
                setRowWidth('160');
                // Set Damages and Clearing rates for 160 ft ROW
                setRates(prev => ({...prev, damages: 10000, clearing: 20000}));
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Medium Running Angle',
                  type4: 'Full Tension Deadend',
                  type5: 'N/A'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 70000,
                  type1Labor: 110000,
                  type2Material: 101000,
                  type2Labor: 162000,
                  type3Material: 172000,
                  type3Labor: 273000,
                  type4Material: 167000,
                  type4Labor: 350000,
                  type5Material: 0,
                  type5Labor: 0
                });
              }
              // Auto-populate structure types for 765 kV Steel Lattice (no default costs yet)
              else if (voltage === '765 kV' && e.target.value === 'Steel Lattice') {
                setRowWidth('200');
                // Set Damages and Clearing rates for 200 ft ROW
                setRates(prev => ({...prev, damages: 30000, clearing: 20000}));
                setStructureNames({
                  type1: 'Tangent',
                  type2: 'Small Running Angle',
                  type3: 'Angle Deadend',
                  type4: 'Full Tension Deadend',
                  type5: 'N/A'
                });
                setCustomMode({
                  type1: false,
                  type2: false,
                  type3: false,
                  type4: false,
                  type5: false
                });
                setStructureCosts({
                  type1Material: 154974,
                  type1Labor: 227748,
                  type2Material: 218312,
                  type2Labor: 329483,
                  type3Material: 428385,
                  type3Labor: 632122,
                  type4Material: 527819,
                  type4Labor: 1134719,
                  type5Material: 0,
                  type5Labor: 0
                });
              }
            }}
          >
            <option value="">Select Structure Family</option>
            <option value="Steel Monopole">Steel Monopole</option>
            <option value="Steel Lattice">Steel Lattice</option>
          </select>
        </div>

        {/* Number of Circuits Selection */}
        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <label className="block text-slate-300 text-sm mb-2 font-semibold">How many circuits are being installed?</label>
          <select
            className="w-full max-w-md px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
            value={numCircuits}
            onChange={e => {
              setNumCircuits(e.target.value);
              // Auto-set 2nd circuit fields to N/A when 1 circuit is selected
              if (e.target.value === '1') {
                setConductor2Spec('N/A');
                setShieldWire2Spec('N/A');
                
                // Auto-populate 1st circuit specs for 69 kV - 1 circuit
                if (voltage === '69 kV') {
                  setConductor1Spec('1-959 ACSS/TW "Suwannee"');
                  setShieldWire1Spec('69/138 kV OPGW');
                }
                // Auto-populate 1st circuit specs for 138 kV - 1 circuit
                else if (voltage === '138 kV') {
                  setConductor1Spec('1-1926 ACSS/TW "Cumberland"');
                  setShieldWire1Spec('69/138 kV OPGW');
                }
                // Auto-populate 1st circuit specs for 345 kV - 1 circuit
                else if (voltage === '345 kV') {
                  setConductor1Spec('2-1926 ACSS/TW "Cumberland"');
                  setShieldWire1Spec('345 kV OPGW');
                }
                // Auto-populate 1st circuit specs for 765 kV - 1 circuit
                else if (voltage === '765 kV') {
                  setConductor1Spec('6-795 ACSR "Drake"');
                  setShieldWire1Spec('765 kV OPGW');
                }
              } else if (e.target.value === '2') {
                // Clear N/A when switching to 2 circuits
                if (conductor2Spec === 'N/A') setConductor2Spec('');
                if (shieldWire2Spec === 'N/A') setShieldWire2Spec('');
                
                // Auto-populate conductor and shield wire specs for 69 kV - 2 circuits
                if (voltage === '69 kV') {
                  setConductor1Spec('1-959 ACSS/TW "Suwannee"');
                  setConductor2Spec('1-959 ACSS/TW "Suwannee"');
                  setShieldWire1Spec('69/138 kV OPGW');
                  setShieldWire2Spec('69/138 kV OPGW');
                }
                // Auto-populate conductor and shield wire specs for 138 kV - 2 circuits
                else if (voltage === '138 kV') {
                  setConductor1Spec('1-1926 ACSS/TW "Cumberland"');
                  setConductor2Spec('1-1926 ACSS/TW "Cumberland"');
                  setShieldWire1Spec('69/138 kV OPGW');
                  setShieldWire2Spec('69/138 kV OPGW');
                }
                // Auto-populate conductor and shield wire specs for 345 kV - 2 circuits
                else if (voltage === '345 kV') {
                  setConductor1Spec('2-1926 ACSS/TW "Cumberland"');
                  setConductor2Spec('2-1926 ACSS/TW "Cumberland"');
                  setShieldWire1Spec('345 kV OPGW');
                  setShieldWire2Spec('345 kV OPGW');
                }
              }
            }}
          >
            <option value="1">1</option>
            <option value="2">2</option>
          </select>
        </div>

        <div className="space-y-6">
          {[1, 2, 3, 4, 5].map(num => {
            const currentName = structureNames[`type${num}`];
            const isCustom = customMode[`type${num}`];
            
            return (
              <div key={num} className="bg-slate-900 rounded-lg p-5 border border-slate-700">
                <div className="mb-4">
                  <label className="block text-slate-300 text-sm mb-2 font-semibold">Structure Type #{num}</label>
                  <select
                    className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
                    value={isCustom ? 'Custom' : currentName}
                    onChange={e => {
                      if (e.target.value === 'Custom') {
                        setCustomMode({...customMode, [`type${num}`]: true});
                        setStructureNames({...structureNames, [`type${num}`]: ''});
                      } else {
                        setCustomMode({...customMode, [`type${num}`]: false});
                        setStructureNames({...structureNames, [`type${num}`]: e.target.value});
                      }
                    }}
                  >
                    <option value="Tangent">Tangent</option>
                    <option value="Small Running Angle">Small Running Angle</option>
                    <option value="Medium Running Angle">Medium Running Angle</option>
                    <option value="Running Angle">Running Angle</option>
                    <option value="Angle Deadend">Angle Deadend</option>
                    <option value="Full Tension Deadend">Full Tension Deadend</option>
                    <option value="Custom">Custom</option>
                    <option value="N/A">N/A</option>
                  </select>
                  {isCustom && (
                    <Input 
                      key={`custom-${num}`}
                      className="mt-2"
                      value={currentName} 
                      onChange={e => setStructureNames({...structureNames, [`type${num}`]: e.target.value})} 
                      placeholder="Enter Name"
                    />
                  )}
                </div>
                {currentName !== 'N/A' && (
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="block text-slate-400 text-sm mb-2">Material Cost ($)</label>
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => setStructureCosts({...structureCosts, [`type${num}Material`]: Math.max(0, structureCosts[`type${num}Material`] - 1000)})}
                          className="bg-slate-700 hover:bg-slate-600 text-white px-3 py-2 rounded"
                        >−</button>
                        <div className="flex items-center flex-1">
                          <span className="text-slate-400 mr-2">$</span>
                          <Input 
                            type="text"
                            value={(structureCosts[`type${num}Material`] * (numCircuits === '2' && currentName.includes('Deadend') ? 2 : 1)).toLocaleString('en-US')}
                            onChange={e => {
                              const val = e.target.value.replace(/,/g, '');
                              if (val === '' || !isNaN(val)) {
                                const baseValue = numCircuits === '2' && currentName.includes('Deadend') ? (parseFloat(val) || 0) / 2 : (parseFloat(val) || 0);
                                setStructureCosts({...structureCosts, [`type${num}Material`]: baseValue});
                              }
                            }}
                            placeholder="0"
                          />
                          {numCircuits === '2' && currentName.includes('Deadend') && (
                            <span className="ml-2 text-xs text-blue-400">(×2)</span>
                          )}
                        </div>
                        <button 
                          onClick={() => setStructureCosts({...structureCosts, [`type${num}Material`]: structureCosts[`type${num}Material`] + 1000})}
                          className="bg-slate-700 hover:bg-slate-600 text-white px-3 py-2 rounded"
                        >+</button>
                      </div>
                    </div>
                    <div>
                      <label className="block text-slate-400 text-sm mb-2">Labor Cost ($)</label>
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => setStructureCosts({...structureCosts, [`type${num}Labor`]: Math.max(0, structureCosts[`type${num}Labor`] - 1000)})}
                          className="bg-slate-700 hover:bg-slate-600 text-white px-3 py-2 rounded"
                        >−</button>
                        <div className="flex items-center flex-1">
                          <span className="text-slate-400 mr-2">$</span>
                          <Input 
                            type="text"
                            value={(structureCosts[`type${num}Labor`] * (numCircuits === '2' && currentName.includes('Deadend') ? 2 : 1)).toLocaleString('en-US')}
                            onChange={e => {
                              const val = e.target.value.replace(/,/g, '');
                              if (val === '' || !isNaN(val)) {
                                const baseValue = numCircuits === '2' && currentName.includes('Deadend') ? (parseFloat(val) || 0) / 2 : (parseFloat(val) || 0);
                                setStructureCosts({...structureCosts, [`type${num}Labor`]: baseValue});
                              }
                            }}
                            placeholder="0"
                          />
                          {numCircuits === '2' && currentName.includes('Deadend') && (
                            <span className="ml-2 text-xs text-blue-400">(×2)</span>
                          )}
                        </div>
                        <button 
                          onClick={() => setStructureCosts({...structureCosts, [`type${num}Labor`]: structureCosts[`type${num}Labor`] + 1000})}
                          className="bg-slate-700 hover:bg-slate-600 text-white px-3 py-2 rounded"
                        >+</button>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <div className="mt-8 flex justify-end">
          <Btn onClick={() => setStep(1)}>Continue</Btn>
        </div>
      </div>
    </div>
  );

  if (step === 1) return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
      <div className="w-full max-w-6xl bg-slate-800 rounded-2xl p-8 border border-slate-700">
        <h1 className="text-3xl font-bold text-white mb-6">Project Links</h1>
        <div className="mb-6 flex justify-between">
          <button onClick={() => setStep(0)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={() => setStep(2)} disabled={!links.length}>Continue</Btn>
        </div>
        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <div className="flex gap-3 mb-2">
            <button onClick={downloadTemplate} className="bg-slate-700 hover:bg-slate-600 text-white font-semibold py-3 px-4 rounded-lg">Download Template</button>
            <label className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-4 rounded-lg cursor-pointer text-center">
              <input type="file" accept=".xlsx,.xls" onChange={uploadLinks} className="hidden" />Upload Links
            </label>
          </div>
        </div>
        {links.length > 0 && (
          <div className="mb-6 overflow-x-auto">
            <table className="w-full text-left">
              <thead><tr className="border-b border-slate-700">
                {['Link', 'Length (Miles)', 
                  ...[1, 2, 3, 4, 5].filter(n => structureNames[`type${n}`] !== 'N/A').map(n => structureNames[`type${n}`]),
                  ''
                ].map(h => <th key={h} className="pb-3 text-slate-300 font-semibold">{h}</th>)}
              </tr></thead>
              <tbody>{links.map(l => (
                <tr key={l.id} className="border-b border-slate-700">
                  <td className="py-3"><Input value={l.link} onChange={e => updateLink(l.id, 'link', e.target.value)} placeholder="Link ID" /></td>
                  <td className="py-3 px-2"><Input type="number" min="0" value={l.length} onChange={e => updateLink(l.id, 'length', e.target.value)} placeholder="0" /></td>
                  {[1, 2, 3, 4, 5].filter(n => structureNames[`type${n}`] !== 'N/A').map(n => <td key={n} className="py-3 px-2"><Input type="number" min="0" step="1" value={l[`structureType${n}`]} onChange={e => updateLink(l.id, `structureType${n}`, e.target.value)} placeholder="0" /></td>)}
                  <td className="py-3 px-2"><button onClick={() => removeLink(l.id)} className="text-red-400 hover:text-red-300 text-sm">Remove</button></td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        )}
        <button onClick={addLink} className="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-4 rounded-lg mb-6">+ Add Link</button>
        <div className="flex justify-between">
          <button onClick={() => setStep(0)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={() => setStep(2)} disabled={!links.length}>Continue</Btn>
        </div>
      </div>
    </div>
  );

  if (step === 2) return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 flex items-center justify-center p-4">
      <div className="w-full max-w-4xl bg-slate-800 rounded-2xl p-8 border border-slate-700">
        <h1 className="text-3xl font-bold text-white mb-6">Setup Line Sub-Categories</h1>
        <div className="mb-6 flex justify-between">
          <button onClick={() => setStep(1)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={initCosts}>Continue</Btn>
        </div>
        <div className="space-y-6">
          {categories.filter(mc => mc.key !== 'substationUtility').map(mc => (
            <div key={mc.key} className="bg-slate-900 rounded-lg p-5 border border-slate-700">
              <h3 className="text-white font-semibold mb-3">{mc.label}</h3>
              {subCategories[mc.key]?.map((sc, i) => {
                const isFixed = mc.fixed && i < mc.fixed.length;
                const methodKey = `${mc.key}_${i}`;
                const method = subCategoryMethods[methodKey] || 'fixed';
                return (
                  <div key={i} className="flex justify-between items-center bg-slate-800 px-3 py-2 rounded mb-2">
                    <div className="flex items-center gap-3 flex-1">
                      <span className="text-slate-300">{sc}</span>
                      <select
                        className="px-2 py-1 bg-slate-800 text-slate-300 text-sm rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-8" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.25rem center", backgroundSize: "1.25em 1.25em"}}
                        value={method}
                        onChange={e => setSubCategoryMethods({...subCategoryMethods, [methodKey]: e.target.value})}
                      >
                        <option value="fixed">Fixed</option>
                        <option value="$/mile">$/mile</option>
                        <option value="%">%</option>
                      </select>
                    </div>
                    {!isFixed && <button onClick={() => removeSubCat(mc.key, i)} className="text-red-400 text-sm">Remove</button>}
                  </div>
                );
              })}
              {editingCat === mc.key ? (
                <div className="flex gap-2 mt-3">
                  <Input value={tempSubCat} onChange={e => setTempSubCat(e.target.value)} onKeyPress={e => e.key === 'Enter' && addSubCat(mc.key)} placeholder="Sub-category" autoFocus />
                  <button onClick={() => addSubCat(mc.key)} className="bg-blue-600 text-white px-4 py-2 rounded">Add</button>
                  <button onClick={() => { setEditingCat(null); setTempSubCat(''); }} className="bg-slate-700 text-white px-4 py-2 rounded">Done</button>
                </div>
              ) : <button onClick={() => setEditingCat(mc.key)} className="text-blue-400 text-sm mt-2">+ Add Sub-Category</button>}
            </div>
          ))}
        </div>
        <div className="mt-8 flex justify-between">
          <button onClick={() => setStep(1)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={initCosts}>Continue</Btn>
        </div>
      </div>
    </div>
  );

  if (step === 3) return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 p-4">
      <div className="w-full max-w-4xl mx-auto bg-slate-800 rounded-2xl p-8 border border-slate-700 my-4">
        <h1 className="text-3xl font-bold text-white mb-6">Link Costs</h1>
        <div className="mb-6 flex justify-between items-center">
          <div className="flex gap-3">
            <button onClick={() => setStep(2)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
            <button 
              onClick={exportRatesTemplate}
              className="bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-4 rounded-lg"
            >
              Export Rates
            </button>
            <label className="bg-purple-600 hover:bg-purple-700 text-white font-semibold py-2 px-4 rounded-lg cursor-pointer flex items-center">
              Import Rates
              <input 
                type="file" 
                accept=".xlsx,.xls"
                className="hidden"
                onChange={e => {
                  if (e.target.files?.[0]) {
                    importRatesTemplate(e.target.files[0]);
                    e.target.value = ''; // Reset input
                  }
                }}
              />
            </label>
          </div>
          <Btn onClick={() => setStep(4)}>Continue</Btn>
        </div>

        {/* Voltage and ROW Width */}
        <div className="mb-6 grid grid-cols-2 gap-4">
          <div className="bg-slate-900 rounded-lg p-5 border border-slate-700">
            <label className="block text-slate-300 text-sm mb-2 font-semibold">Voltage</label>
            <div className="relative">
              <select
                className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10"
                style={{
                  backgroundImage: "url(\"data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%239ca3af' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e\")",
                  backgroundPosition: "right 0.5rem center",
                  backgroundRepeat: "no-repeat",
                  backgroundSize: "1.5em 1.5em"
                }}
                value={voltage}
                onChange={e => {
                setVoltage(e.target.value);
                // Auto-assign ROW Width based on voltage and structure family
                const v = e.target.value;
                if (v === '69 kV' && structureFamily === 'Steel Monopole') {
                  setRowWidth('70');
                } else if (v === '138 kV' && structureFamily === 'Steel Monopole') {
                  setRowWidth('70');
                } else if (v === '345 kV' && structureFamily === 'Steel Monopole') {
                  setRowWidth('100');
                } else if (v === '345 kV' && structureFamily === 'Steel Lattice') {
                  setRowWidth('160');
                } else if (v === '765 kV' && structureFamily === 'Steel Lattice') {
                  setRowWidth('200');
                }
              }}
            >
              <option value="">Select Voltage</option>
              <option value="69 kV">69 kV</option>
              <option value="138 kV">138 kV</option>
              <option value="345 kV">345 kV</option>
              <option value="765 kV">765 kV</option>
            </select>
            </div>
          </div>
          <div className="bg-slate-900 rounded-lg p-5 border border-slate-700">
            <label className="block text-slate-300 text-sm mb-2 font-semibold">ROW Width (ft.)</label>
            <select
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
              value={rowWidth}
              onChange={e => {
                setRowWidth(e.target.value);
                
                // Set default Damages rate based on ROW Width
                const damagesRates = {
                  '70': 6500,
                  '100': 6500,
                  '160': 10000,
                  '200': 30000
                };
                // Set default Clearing rate based on ROW Width
                const clearingRates = {
                  '70': 15000,
                  '100': 15000,
                  '160': 20000,
                  '200': 20000
                };
                if (damagesRates[e.target.value]) {
                  setRates({
                    ...rates, 
                    damages: damagesRates[e.target.value],
                    clearing: clearingRates[e.target.value]
                  });
                }
              }}
            >
              <option value="">Select ROW Width</option>
              <option value="70">70</option>
              <option value="100">100</option>
              <option value="160">160</option>
              <option value="200">200</option>
            </select>
          </div>
        </div>

        {/* Conductor and Shield Wire Specifications */}
        <div className="mb-6 grid grid-cols-2 gap-4">
          <div className="bg-slate-900 rounded-lg p-5 border border-slate-700">
            <label className="block text-slate-300 text-sm mb-2 font-semibold">1st Circuit Conductor</label>
            <select
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
              value={conductor1Spec}
              onChange={e => {
                setConductor1Spec(e.target.value);
                
                // Set default conductor material and labor rates based on specification
                const conductorRates = {
                  '1-795 ACSR "Drake"': { material: 14698, labor: 15258 },
                  '1-959 ACSS/TW "Suwannee"': { material: 17711, labor: 19645 },
                  '1-1926 ACSS/TW "Cumberland"': { material: 34952, labor: 28842 },
                  '2-959 ACSS/TW "Suwannee"': { material: 35421, labor: 28985 },
                  '2-1926 ACSS/TW "Cumberland"': { material: 69890, labor: 54111 },
                  '6-795 ACSR "Drake"': { material: 88188, labor: 87018 }
                };
                
                if (conductorRates[e.target.value]) {
                  setRates({
                    ...rates,
                    conductor1Material: conductorRates[e.target.value].material,
                    wireLabor: conductorRates[e.target.value].labor
                  });
                }
              }}
            >
              <option value="">Select Conductor</option>
              <option value='1-795 ACSR "Drake"'>1-795 ACSR "Drake"</option>
              <option value='1-959 ACSS/TW "Suwannee"'>1-959 ACSS/TW "Suwannee"</option>
              <option value='1-1926 ACSS/TW "Cumberland"'>1-1926 ACSS/TW "Cumberland"</option>
              <option value='2-959 ACSS/TW "Suwannee"'>2-959 ACSS/TW "Suwannee"</option>
              <option value='2-1926 ACSS/TW "Cumberland"'>2-1926 ACSS/TW "Cumberland"</option>
              <option value='6-795 ACSR "Drake"'>6-795 ACSR "Drake"</option>
              <option value="N/A">N/A</option>
            </select>
          </div>
          <div className="bg-slate-900 rounded-lg p-5 border border-slate-700">
            <label className="block text-slate-300 text-sm mb-2 font-semibold">2nd Circuit Conductor</label>
            <select
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
              value={conductor2Spec}
              onChange={e => {
                setConductor2Spec(e.target.value);
                
                // Set default conductor material and labor rates based on specification
                const conductorRates = {
                  '1-795 ACSR "Drake"': { material: 14698, labor: 15258 },
                  '1-959 ACSS/TW "Suwannee"': { material: 17711, labor: 19645 },
                  '1-1926 ACSS/TW "Cumberland"': { material: 34952, labor: 28842 },
                  '2-959 ACSS/TW "Suwannee"': { material: 35421, labor: 28985 },
                  '2-1926 ACSS/TW "Cumberland"': { material: 69890, labor: 54111 },
                  '6-795 ACSR "Drake"': { material: 88188, labor: 87018 }
                };
                
                if (conductorRates[e.target.value]) {
                  setRates({
                    ...rates,
                    conductor2Material: conductorRates[e.target.value].material,
                    conductor2Labor: conductorRates[e.target.value].labor
                  });
                }
              }}
            >
              <option value="">Select Conductor</option>
              <option value='1-795 ACSR "Drake"'>1-795 ACSR "Drake"</option>
              <option value='1-959 ACSS/TW "Suwannee"'>1-959 ACSS/TW "Suwannee"</option>
              <option value='1-1926 ACSS/TW "Cumberland"'>1-1926 ACSS/TW "Cumberland"</option>
              <option value='2-959 ACSS/TW "Suwannee"'>2-959 ACSS/TW "Suwannee"</option>
              <option value='2-1926 ACSS/TW "Cumberland"'>2-1926 ACSS/TW "Cumberland"</option>
              <option value='6-795 ACSR "Drake"'>6-795 ACSR "Drake"</option>
              <option value="N/A">N/A</option>
            </select>
          </div>
        </div>

        <div className="mb-6 grid grid-cols-2 gap-4">
          <div className="bg-slate-900 rounded-lg p-5 border border-slate-700">
            <label className="block text-slate-300 text-sm mb-2 font-semibold">1st Circuit Shield Wire</label>
            <select
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
              value={shieldWire1Spec}
              onChange={e => {
                setShieldWire1Spec(e.target.value);
                
                // Set default shield wire material and labor rates based on specification
                const shieldWireRates = {
                  '3/8" EHS Steel': { material: 1958, labor: 5088 },
                  '7/16" EHS Steel': { material: 3753, labor: 5088 },
                  '69/138 kV OPGW': { material: 6635, labor: 12705 },
                  '345 kV OPGW': { material: 19035, labor: 13977 },
                  '765 kV OPGW': { material: 19035, labor: 13977 }
                };
                
                if (shieldWireRates[e.target.value]) {
                  setRates({
                    ...rates,
                    shieldWire1Material: shieldWireRates[e.target.value].material,
                    shieldWire1Labor: shieldWireRates[e.target.value].labor
                  });
                }
              }}
            >
              <option value="">Select Shield Wire</option>
              <option value='3/8" EHS Steel'>3/8" EHS Steel</option>
              <option value='7/16" EHS Steel'>7/16" EHS Steel</option>
              <option value="69/138 kV OPGW">69/138 kV OPGW</option>
              <option value="345 kV OPGW">345 kV OPGW</option>
              <option value="765 kV OPGW">765 kV OPGW</option>
              <option value="N/A">N/A</option>
            </select>
          </div>
          <div className="bg-slate-900 rounded-lg p-5 border border-slate-700">
            <label className="block text-slate-300 text-sm mb-2 font-semibold">2nd Circuit Shield Wire</label>
            <select
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
              value={shieldWire2Spec}
              onChange={e => {
                setShieldWire2Spec(e.target.value);
                
                // Set default shield wire material and labor rates based on specification
                const shieldWireRates = {
                  '3/8" EHS Steel': { material: 1958, labor: 5088 },
                  '7/16" EHS Steel': { material: 3753, labor: 5088 },
                  '69/138 kV OPGW': { material: 6635, labor: 12705 },
                  '345 kV OPGW': { material: 19035, labor: 13977 },
                  '765 kV OPGW': { material: 19035, labor: 13977 }
                };
                
                if (shieldWireRates[e.target.value]) {
                  setRates({
                    ...rates,
                    shieldWire2Material: shieldWireRates[e.target.value].material,
                    shieldWire2Labor: shieldWireRates[e.target.value].labor
                  });
                }
              }}
            >
              <option value="">Select Shield Wire</option>
              <option value='3/8" EHS Steel'>3/8" EHS Steel</option>
              <option value='7/16" EHS Steel'>7/16" EHS Steel</option>
              <option value="69/138 kV OPGW">69/138 kV OPGW</option>
              <option value="345 kV OPGW">345 kV OPGW</option>
              <option value="765 kV OPGW">765 kV OPGW</option>
              <option value="N/A">N/A</option>
            </select>
          </div>
        </div>

        <div className="space-y-6">
          {categories.map(mc => {
            if (!subCategories[mc.key]) return null;
            return (
              <div key={mc.key} className="bg-slate-900 rounded-lg p-5 border border-slate-700">
                <h3 className="text-white font-semibold mb-4">{mc.label}</h3>
                <div className="space-y-4">
                  {subCategories[mc.key].map((sc, si) => {
                    // Skip Land Acquisition (index 0 in rightOfWay) - it's now route-based
                    if (mc.key === 'rightOfWay' && si === 0) return null;
                    
                    const methodKey = `${mc.key}_${si}`;
                    const method = subCategoryMethods[methodKey] || 'fixed';
                    const isStructMat = mc.key === 'procurement' && si === 0;
                    const isStructLab = mc.key === 'transmissionContract' && si === 5;
                    const isTax = mc.key === 'procurement' && si === 5;
                    const isStor = mc.key === 'procurement' && si === 6;
                    const isOH = mc.key === 'transmissionContract' && si === 10;
                    const isConductor1 = mc.key === 'procurement' && si === 1;
                    const isConductor2 = mc.key === 'procurement' && si === 2;
                    const isShieldWire1 = mc.key === 'procurement' && si === 3;
                    const isShieldWire2 = mc.key === 'procurement' && si === 4;
                    
                    // Auto if: structure material/labor, tax, stores, overhead, OR if method is $/mile or %
                    const auto = isStructMat || isStructLab || isTax || isStor || isOH || method === '$/mile' || method === '%';
                    
                    const rateConfigs = [
                      { cond: mc.key === 'rightOfWay' && si === 1, type: 'rowAgent', label: 'ROW Agent Rate ($/mi)' },
                      { cond: mc.key === 'engineeringContract' && si === 0, type: 'surveying', label: 'Land Surveying and SUE Rate ($/mi)' },
                      { cond: mc.key === 'engineeringContract' && si === 1, type: 'aerialSurvey', label: 'Aerial Survey Rate ($/mi)' },
                      { cond: mc.key === 'engineeringContract' && si === 2, type: 'soilBorings', label: 'Soil Borings Rate ($/mi)' },
                      { cond: mc.key === 'engineeringContract' && si === 3, type: 'engineering', label: 'Engineering and Design Rate ($/mi)' },
                      { cond: isConductor1 && method === '$/mile', type: 'conductor1Material', label: '1st Circuit Conductor Material Rate ($/mi)' },
                      { cond: isConductor2 && method === '$/mile', type: 'conductor2Material', label: '2nd Circuit Conductor Material Rate ($/mi)' },
                      { cond: isShieldWire1 && method === '$/mile', type: 'shieldWire1Material', label: '1st Circuit Shield Wire Material Rate ($/mi)' },
                      { cond: isShieldWire2 && method === '$/mile', type: 'shieldWire2Material', label: '2nd Circuit Shield Wire Material Rate ($/mi)' },
                      { cond: isTax, type: 'tax', label: 'M&S Rate (%)' },
                      { cond: isStor, type: 'stores', label: 'P&S Rate (%)' },
                      { cond: mc.key === 'transmissionContract' && si === 0, type: 'clearing', label: 'Clearing Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 1, type: 'environmental', label: 'Environmental Controls Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 2, type: 'damages', label: 'Damages Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 3, type: 'constructionMgmt', label: 'Construction Management Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 4, type: 'permitting', label: 'Permitting Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 6, type: 'wireLabor', label: '1st Circuit Conductor Labor Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 7, type: 'conductor2Labor', label: '2nd Circuit Conductor Labor Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 8, type: 'shieldWire1Labor', label: '1st Circuit Shield Wire Labor Rate ($/mi)' },
                      { cond: mc.key === 'transmissionContract' && si === 9, type: 'shieldWire2Labor', label: '2nd Circuit Shield Wire Labor Rate ($/mi)' },
                      { cond: isOH, type: 'overhead', label: 'Construction Overhead Rate (%)' }
                    ];
                    const rateConfig = rateConfigs.find(r => r.cond);
                    
                    return (
                      <div key={si} className="bg-slate-800 rounded p-4">
                        <h4 className="text-slate-300 font-semibold mb-3">
                          {sc} 
                          {auto && <span className="text-xs text-blue-400"> (auto)</span>}
                          {!auto && subCategoryMethods[`${mc.key}_${si}`] === 'fixed' && <span className="text-xs text-blue-400"> (fixed)</span>}
                        </h4>
                        {rateConfig && (
                          <div className="mb-3 bg-slate-900 p-3 rounded">
                            <label className="block text-slate-400 text-sm mb-2">{rateConfig.label}</label>
                            {/* Show selected conductor/shield wire spec */}
                            {isConductor1 && conductor1Spec && conductor1Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{conductor1Spec}</div>
                            )}
                            {isConductor2 && conductor2Spec && conductor2Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{conductor2Spec}</div>
                            )}
                            {isShieldWire1 && shieldWire1Spec && shieldWire1Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{shieldWire1Spec}</div>
                            )}
                            {isShieldWire2 && shieldWire2Spec && shieldWire2Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{shieldWire2Spec}</div>
                            )}
                            {/* Show spec for conductor labor rates */}
                            {mc.key === 'transmissionContract' && si === 6 && conductor1Spec && conductor1Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{conductor1Spec}</div>
                            )}
                            {mc.key === 'transmissionContract' && si === 7 && conductor2Spec && conductor2Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{conductor2Spec}</div>
                            )}
                            {/* Show spec for shield wire labor rates */}
                            {mc.key === 'transmissionContract' && si === 8 && shieldWire1Spec && shieldWire1Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{shieldWire1Spec}</div>
                            )}
                            {mc.key === 'transmissionContract' && si === 9 && shieldWire2Spec && shieldWire2Spec !== 'N/A' && (
                              <div className="text-xs text-slate-500 mb-2">{shieldWire2Spec}</div>
                            )}
                            <Input type="number" step="1000" value={rates[rateConfig.type]} onChange={e => updateRate(rateConfig.type, parseFloat(e.target.value) || 0)} className="w-32" />
                          </div>
                        )}
                        <div className="space-y-2">
                          {links.map(l => {
                            const k = `${mc.key}_${si}_${l.id}`;
                            const calcVal = (t) => ((parseFloat(l.length) || 0) * rates[t]).toFixed(2);
                            const calcStructMat = () => {
                              const st1 = parseFloat(l.structureType1) || 0;
                              const st2 = parseFloat(l.structureType2) || 0;
                              const st3 = parseFloat(l.structureType3) || 0;
                              const st4 = parseFloat(l.structureType4) || 0;
                              const st5 = parseFloat(l.structureType5) || 0;
                              
                              // Apply circuit multiplier for Deadend structures
                              const multiplier = numCircuits === '2' ? 2 : 1;
                              const cost1 = structureNames.type1.includes('Deadend') ? structureCosts.type1Material * multiplier : structureCosts.type1Material;
                              const cost2 = structureNames.type2.includes('Deadend') ? structureCosts.type2Material * multiplier : structureCosts.type2Material;
                              const cost3 = structureNames.type3.includes('Deadend') ? structureCosts.type3Material * multiplier : structureCosts.type3Material;
                              const cost4 = structureNames.type4.includes('Deadend') ? structureCosts.type4Material * multiplier : structureCosts.type4Material;
                              const cost5 = structureNames.type5.includes('Deadend') ? structureCosts.type5Material * multiplier : structureCosts.type5Material;
                              
                              return (
                                (st1 * cost1) +
                                (st2 * cost2) +
                                (st3 * cost3) +
                                (st4 * cost4) +
                                (st5 * cost5)
                              ).toFixed(2);
                            };
                            const calcStructLab = () => {
                              const st1 = parseFloat(l.structureType1) || 0;
                              const st2 = parseFloat(l.structureType2) || 0;
                              const st3 = parseFloat(l.structureType3) || 0;
                              const st4 = parseFloat(l.structureType4) || 0;
                              const st5 = parseFloat(l.structureType5) || 0;
                              
                              // Apply circuit multiplier for Deadend structures
                              const multiplier = numCircuits === '2' ? 2 : 1;
                              const cost1 = structureNames.type1.includes('Deadend') ? structureCosts.type1Labor * multiplier : structureCosts.type1Labor;
                              const cost2 = structureNames.type2.includes('Deadend') ? structureCosts.type2Labor * multiplier : structureCosts.type2Labor;
                              const cost3 = structureNames.type3.includes('Deadend') ? structureCosts.type3Labor * multiplier : structureCosts.type3Labor;
                              const cost4 = structureNames.type4.includes('Deadend') ? structureCosts.type4Labor * multiplier : structureCosts.type4Labor;
                              const cost5 = structureNames.type5.includes('Deadend') ? structureCosts.type5Labor * multiplier : structureCosts.type5Labor;
                              
                              return (
                                (st1 * cost1) +
                                (st2 * cost2) +
                                (st3 * cost3) +
                                (st4 * cost4) +
                                (st5 * cost5)
                              ).toFixed(2);
                            };
                            const calcOH = () => {
                              let base = 0;
                              for (let i = 0; i < 10; i++) base += parseFloat(inputs[`transmissionContract_${i}_${l.id}`]) || 0;
                              return (base * rates.overhead / 100).toFixed(2);
                            };
                            
                            // Determine value based on method and special cases
                            let v = inputs[k] || '';
                            if (isStructMat) v = calcStructMat();
                            else if (isStructLab) v = calcStructLab();
                            else if (isOH) v = calcOH();
                            else if (method === '$/mile' && rateConfig) v = calcVal(rateConfig.type);
                            else v = inputs[k] || '';
                            
                            return (
                              <div key={k} className="flex items-center gap-3">
                                <span className="text-slate-400 text-sm w-32 truncate">{l.link}</span>
                                {(isStructMat || isStructLab) ? (
                                  <div className="text-slate-400 text-xs w-auto flex-1">
                                    <div className="mb-1">
                                      {[l.structureType1, l.structureType2, l.structureType3, l.structureType4, l.structureType5].map(s => parseFloat(s) || 0).reduce((a, b) => a + b, 0)} total qty
                                    </div>
                                    {[1, 2, 3, 4, 5].filter(n => structureNames[`type${n}`] !== 'N/A' && parseFloat(l[`structureType${n}`] || 0) > 0).map(n => (
                                      <div key={n} className="text-slate-500 text-xs pl-2">
                                        {structureNames[`type${n}`]}: {l[`structureType${n}`] || 0}
                                      </div>
                                    ))}
                                  </div>
                                ) : method === '$/mile' ? (
                                  <span className="text-slate-500 text-xs w-20">{l.length} mi</span>
                                ) : (
                                  <span className="w-20"></span>
                                )}
                                <div className="flex items-center flex-1 min-w-0">
                                  <span className="text-slate-400 mr-2">$</span>
                                  <Input type="number" step="0.01" value={v} onChange={e => updateInput(k, e.target.value)} disabled={auto} className={`text-right ${auto ? 'opacity-60' : ''}`} />
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
        <div className="mt-8 flex justify-between">
          <button onClick={() => setStep(2)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={() => setStep(4)}>Continue</Btn>
        </div>
      </div>
    </div>
  );

  if (step === 4) return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 p-4">
      <div className="w-full max-w-6xl mx-auto bg-slate-800 rounded-2xl p-8 border border-slate-700 my-4">
        <h1 className="text-3xl font-bold text-white mb-6">Link Cost Summary</h1>
        
        <div className="mb-6 flex justify-between">
          <button onClick={() => setStep(3)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={() => setStep(5)}>Continue</Btn>
        </div>

        <div className="mb-6">
          <button onClick={exportLinkSummary} className="bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-6 rounded-lg">
            Export Link Costs Summary
          </button>
        </div>

        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <label className="block text-slate-300 font-semibold mb-3">Contingency (%) - Applied to all categories</label>
          <div className="flex items-center gap-3 max-w-sm">
            <input
              type="number" 
              step="0.1"
              min="0"
              max="100"
              value={tempContingency}
              onInput={e => {
                setTempContingency(e.target.value);
                setContingencyChanged(true);
              }}
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 text-right w-32" style={{WebkitAppearance: "none", MozAppearance: "none"}}
              placeholder="0"
            />
            <span className="text-slate-400">%</span>
            {contingencyChanged && (
              <button 
                onClick={() => {
                  setContingency(parseFloat(tempContingency) || 0);
                  setContingencyChanged(false);
                }}
                className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg flex items-center gap-2"
                title="Update costs"
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M21.5 2v6h-6M2.5 22v-6h6M2 11.5a10 10 0 0 1 18.8-4.3M22 12.5a10 10 0 0 1-18.8 4.2"/>
                </svg>
                Update
              </button>
            )}
          </div>
        </div>

        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <label className="block text-slate-300 font-semibold mb-3">Escalation - Compounded annually, applied to all categories</label>
          <div className="grid grid-cols-3 gap-4 max-w-2xl">
            <div>
              <label className="block text-slate-400 text-sm mb-2">Annual Rate (%)</label>
              <div className="flex items-center">
                <Input 
                  type="number" 
                  step="0.1"
                  min="0"
                  max="100"
                  value={escalationRate}
                  onChange={e => setEscalationRate(parseFloat(e.target.value) || 0)}
                  className="text-right w-32"
                  placeholder="0"
                />
                <span className="text-slate-400 ml-2">%</span>
              </div>
            </div>
            <div>
              <label className="block text-slate-400 text-sm mb-2">Years</label>
              <Input 
                type="number" 
                min="0"
                value={escalationYears}
                onChange={e => setEscalationYears(parseInt(e.target.value) || 0)}
                className="text-right w-32"
                placeholder="0"
              />
            </div>
            <div>
              <label className="block text-slate-400 text-sm mb-2">Months</label>
              <Input 
                type="number" 
                min="0"
                max="11"
                value={escalationMonths}
                onChange={e => setEscalationMonths(parseInt(e.target.value) || 0)}
                className="text-right w-32"
                placeholder="0"
              />
            </div>
          </div>
        </div>

        <div className="space-y-6">
          {links.map(link => {
            // Calculate base category totals
            const linkCategoryTotals = {};
            categories.slice(0, 7).forEach(mc => {
              if (subCategories[mc.key]) {
                let categoryTotal = 0;
                subCategories[mc.key].forEach((sc, si) => {
                  const key = `${mc.key}_${si}_${link.id}`;
                  categoryTotal += round(inputs[key]);
                });
                linkCategoryTotals[mc.key] = categoryTotal;
              }
            });
            
            // Calculate escalation and contingency factors
            const totalYears = escalationYears + (escalationMonths / 12);
            const escalationFactor = Math.pow(1 + (escalationRate / 100), totalYears);
            const contingencyFactor = 1 + (contingency / 100);
            
            // Apply escalation and contingency to each category, then round up to nearest 1,000
            const linkCategoryTotalsAdjusted = {};
            categories.slice(0, 7).forEach(mc => {
              const baseValue = linkCategoryTotals[mc.key] || 0;
              
              // Apply escalation
              const afterEscalation = baseValue * escalationFactor;
              
              // Apply contingency
              const afterContingency = afterEscalation * contingencyFactor;
              
              // Round final category value up to nearest 1,000
              linkCategoryTotalsAdjusted[mc.key] = Math.ceil(afterContingency / 1000) * 1000;
            });
            
            // Calculate link total from rounded category values
            const linkTotal = Object.values(linkCategoryTotalsAdjusted).reduce((sum, val) => sum + val, 0);
            
            return (
              <div key={link.id} className="bg-slate-900 rounded-lg p-5 border border-slate-700">
                <div className="mb-4">
                  <h3 className="text-white font-semibold text-lg">Link {link.link}</h3>
                  <span className="text-slate-400 text-sm">{link.length} miles</span>
                  <div className="text-slate-400 text-xs mt-1">
                    {(() => {
                      // Calculate structure counts for this link
                      const structureCounts = {};
                      [1, 2, 3, 4, 5].forEach(n => {
                        const typeName = structureNames[`type${n}`];
                        if (typeName && typeName !== 'N/A') {
                          const count = parseFloat(link[`structureType${n}`]) || 0;
                          if (count > 0) {
                            structureCounts[typeName] = count;
                          }
                        }
                      });
                      
                      // Format structure counts
                      const structureText = Object.entries(structureCounts)
                        .map(([name, count]) => `${name}: ${count}`)
                        .join(', ');
                      
                      return structureText ? `Str. Count: ${structureText}` : 'No structures';
                    })()}
                  </div>
                </div>

                <div className="space-y-1">
                  {categories.slice(0, 7).map(mc => {
                    const categoryTotal = linkCategoryTotalsAdjusted[mc.key] || 0;
                    
                    return (
                      <div key={mc.key} className="flex justify-between text-sm">
                        <span className="text-slate-400">{mc.label}</span>
                        <span className="text-slate-300 text-right">${categoryTotal.toLocaleString('en-US')}</span>
                      </div>
                    );
                  })}
                  
                  <div className="flex justify-between text-sm pt-2 mt-2 border-t border-slate-700">
                    <span className="text-white font-bold">Estimated Total Cost</span>
                    <span className="text-white font-bold text-right">${linkTotal.toLocaleString('en-US')}</span>
                  </div>
                </div>
              </div>
            );
          })}
        </div>

        <div className="mt-8 flex justify-between">
          <button onClick={() => setStep(3)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={() => setStep(5)}>Continue</Btn>
        </div>
      </div>
    </div>
  );

  if (step === 5) return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 p-4">
      <div className="w-full max-w-4xl mx-auto bg-slate-800 rounded-2xl p-8 border border-slate-700 my-4">
        <h1 className="text-3xl font-bold text-white mb-6">Create Routes</h1>
        
        <div className="mb-6 flex justify-between">
          <button onClick={() => setStep(4)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <button onClick={() => setStep(6)} className="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-8 rounded-lg">View Summary</button>
        </div>

        <p className="text-slate-300 mb-6">Create routes and select which links belong to each route.</p>
        
        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <label className="block text-slate-300 font-semibold mb-3">Estimated Substation Facilities Cost (applies to all routes)</label>
          <div className="flex items-center max-w-sm">
            <span className="text-slate-400 mr-2">$</span>
            <input
              type="number" 
              step="1000"
              value={substationCost}
              onInput={e => setSubstationCost(e.target.value)}
              className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 text-right" style={{WebkitAppearance: "none", MozAppearance: "none"}}
              placeholder="0"
            />
          </div>
        </div>
        
        <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
          <div className="flex gap-3">
            <button 
              onClick={() => {
                const ws = XLSX.utils.aoa_to_sheet([
                  ['Route ID', 'Route Link Combination (dash-separated)', 'Land Acquisition Cost'],
                  ['1', links.map(l => l.link).slice(0, 2).join('-'), '0'],
                  ['', '', '']
                ]);
                ws['!cols'] = [{ wch: 25 }, { wch: 40 }, { wch: 25 }];
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Routes');
                XLSX.writeFile(wb, 'Routes_Template.xlsx');
              }}
              className="bg-slate-700 hover:bg-slate-600 text-white font-semibold py-3 px-4 rounded-lg"
            >
              Download Template
            </button>
            <label className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 px-4 rounded-lg cursor-pointer text-center">
              <input 
                type="file" 
                accept=".xlsx,.xls" 
                onChange={e => {
                  const file = e.target.files[0];
                  if (!file) return;
                  const reader = new FileReader();
                  reader.onload = (ev) => {
                    try {
                      const wb = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
                      const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
                      
                      const newRoutes = data.slice(1).filter(r => r[0] && r[0].toString().trim()).map(row => {
                        const routeName = row[0].toString().trim();
                        const linksStr = row[1] ? row[1].toString() : '';
                        const landAcqCost = row[2] ? parseFloat(row[2]) || 0 : 0;
                        const linkNames = linksStr.split('-').map(s => s.trim()).filter(s => s);
                        
                        const linkIds = [];
                        linkNames.forEach(linkName => {
                          const link = links.find(l => l.link === linkName);
                          if (link) {
                            linkIds.push(link.id);
                          }
                        });
                        
                        const routeId = Date.now() + Math.random();
                        // Set land acquisition cost for this route
                        setRouteLandAcquisition(prev => ({...prev, [routeId]: landAcqCost}));
                        
                        return { id: routeId, name: routeName, linkIds };
                      });
                      
                      setRoutes(newRoutes);
                      e.target.value = '';
                    } catch (err) {
                      alert('Error reading file. Please use the template format.');
                    }
                  };
                  reader.readAsArrayBuffer(file);
                }}
                className="hidden" 
              />
              Upload Route Link Combo
            </label>
          </div>
          <p className="text-slate-400 text-xs mt-2">Template: List Route ID and Link Names separated by dashes (e.g., "A-B-C").</p>
        </div>
        
        <div className="mb-6">
          <button 
            onClick={() => setRoutes([...routes, { id: Date.now(), name: `${routes.length + 1}`, linkIds: [] }])}
            className="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-4 rounded-lg"
          >
            + Add New Route
          </button>
        </div>

        {routes.length > 0 ? (
          <div className="space-y-4">
            {routes.map((route, idx) => {
              const routeLinks = links.filter(l => route.linkIds.includes(l.id));
              const totalLength = routeLinks.reduce((sum, l) => sum + (parseFloat(l.length) || 0), 0);
              
              return (
                <div key={route.id} className="bg-slate-900 rounded-lg p-5 border border-slate-700">
                  <div className="flex justify-between items-start mb-4">
                    <div className="flex-1 mr-4">
                      <label className="block text-slate-300 text-sm mb-2">Route ID</label>
                      <Input 
                        value={route.name}
                        onChange={e => {
                          const newRoutes = [...routes];
                          newRoutes[idx].name = e.target.value;
                          setRoutes(newRoutes);
                        }}
                        placeholder={`${idx + 1}`}
                      />
                    </div>
                    <button 
                      onClick={() => {
                        if (confirm(`Remove Route ${route.name || idx + 1}?`)) {
                          setRoutes(routes.filter((_, i) => i !== idx));
                        }
                      }}
                      className="text-red-400 hover:text-red-300 text-sm mt-6"
                    >
                      Remove Route
                    </button>
                  </div>

                  <div className="mb-3">
                    <label className="block text-slate-300 text-sm mb-2">Select Links for this Route</label>
                    <div className="mb-3 flex gap-2">
                      <input
                        key={`input-${route.id}`}
                        defaultValue={linkInputs[route.id] || ''}
                        className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 pr-10" style={{backgroundImage: "url(\"data:image/svg+xml,%3Csvg xmlns=\"http://www.w3.org/2000/svg\" fill=\"none\" viewBox=\"0 0 24 24\" stroke=\"%239ca3af\"%3E%3Cpath stroke-linecap=\"round\" stroke-linejoin=\"round\" stroke-width=\"2\" d=\"M19 9l-7 7-7-7\" /%3E%3C/svg%3E\")", backgroundRepeat: "no-repeat", backgroundPosition: "right 0.5rem center", backgroundSize: "1.5em 1.5em"}}
                        placeholder="Enter links separated by dashes (e.g., A-B-C)"
                        onKeyPress={e => {
                          if (e.key === 'Enter') {
                            const value = e.target.value.trim();
                            if (value) {
                              const linkNames = value.split('-').map(s => s.trim()).filter(s => s);
                              const newLinkIds = [];
                              linkNames.forEach(linkName => {
                                const link = links.find(l => l.link === linkName);
                                if (link) {
                                  newLinkIds.push(link.id);
                                }
                              });
                              const newRoutes = [...routes];
                              newRoutes[idx].linkIds = newLinkIds;
                              setRoutes(newRoutes);
                              setLinkInputs({...linkInputs, [route.id]: value});
                            }
                          }
                        }}
                      />
                      <button
                        onClick={(e) => {
                          const input = e.target.parentElement.querySelector('input');
                          const value = input.value.trim();
                          if (value) {
                            const linkNames = value.split('-').map(s => s.trim()).filter(s => s);
                            const newLinkIds = [];
                            linkNames.forEach(linkName => {
                              const link = links.find(l => l.link === linkName);
                              if (link) {
                                newLinkIds.push(link.id);
                              }
                            });
                            const newRoutes = [...routes];
                            newRoutes[idx].linkIds = newLinkIds;
                            setRoutes(newRoutes);
                            setLinkInputs({...linkInputs, [route.id]: value});
                          }
                        }}
                        className="bg-green-600 hover:bg-green-700 text-white font-semibold px-4 py-2 rounded-lg whitespace-nowrap"
                      >
                        Update
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {links.map(l => {
                        const isSelected = route.linkIds.includes(l.id);
                        return (
                          <button
                            key={l.id}
                            onClick={(e) => {
                              const newRoutes = [...routes];
                              if (isSelected) {
                                newRoutes[idx].linkIds = route.linkIds.filter(id => id !== l.id);
                              } else {
                                newRoutes[idx].linkIds = [...route.linkIds, l.id];
                              }
                              // Update the input field value
                              const selectedLinks = links.filter(link => newRoutes[idx].linkIds.includes(link.id));
                              const newValue = selectedLinks.map(link => link.link).join('-');
                              
                              // Find the input element and update its value
                              const routeCard = e.target.closest('.bg-slate-900');
                              const input = routeCard.querySelector('input');
                              if (input) {
                                input.value = newValue;
                              }
                              
                              setLinkInputs({...linkInputs, [route.id]: newValue});
                              setRoutes(newRoutes);
                            }}
                            className={`px-3 py-2 rounded text-sm ${
                              isSelected 
                                ? 'bg-blue-600 text-white' 
                                : 'bg-slate-800 text-slate-300 hover:bg-slate-700'
                            }`}
                          >
                            {l.link}
                          </button>
                        );
                      })}
                    </div>
                  </div>

                  {route.linkIds.length > 0 && (
                    <div className="mt-3 pt-3 border-t border-slate-700">
                      <div className="flex justify-between mb-3">
                        <span className="text-slate-400 text-sm">Total Length:</span>
                        <span className="text-white font-semibold">{totalLength.toFixed(2)} miles</span>
                      </div>
                      <div className="mt-3">
                        <label className="block text-slate-300 text-sm mb-2">Land Acquisition Cost (fixed for this route)</label>
                        <div className="flex items-center">
                          <span className="text-slate-400 mr-2">$</span>
                          <input
                            type="number"
                            step="1000"
                            value={routeLandAcquisition[route.id] || ''}
                            onInput={e => setRouteLandAcquisition({...routeLandAcquisition, [route.id]: e.target.value})}
                            className="w-full px-3 py-2 bg-slate-800 text-white rounded border border-slate-600 focus:outline-none focus:border-blue-500 text-right" style={{WebkitAppearance: "none", MozAppearance: "none"}}
                            placeholder="0"
                          />
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        ) : (
          <div className="bg-slate-900 rounded-lg p-8 border border-slate-700 text-center">
            <p className="text-slate-400">No routes created yet. Click "Add New Route" above to get started.</p>
          </div>
        )}

        <div className="mt-8 mb-4">
          <button 
            onClick={() => setRoutes([...routes, { id: Date.now(), name: `${routes.length + 1}`, linkIds: [] }])}
            className="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-4 rounded-lg"
          >
            + Add New Route
          </button>
        </div>

        <div className="mt-8 flex justify-between">
          <button onClick={() => setStep(4)} className="bg-slate-700 text-white font-bold py-3 px-8 rounded-lg">Back</button>
          <Btn onClick={() => setStep(6)}>View Summary</Btn>
        </div>
      </div>
    </div>
  );

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-blue-900 to-slate-900 p-4">
      <div className="w-full max-w-4xl mx-auto bg-slate-800 rounded-2xl p-8 border border-slate-700 my-4">
        <h1 className="text-3xl font-bold text-white mb-6">Route Summary</h1>
        
        <div className="mb-6 flex flex-wrap gap-3">
          <button onClick={() => setStep(0)} className="text-slate-400 hover:text-white text-sm underline">Edit Structures</button>
          <button onClick={() => setStep(1)} className="text-slate-400 hover:text-white text-sm underline">Edit Links</button>
          <button onClick={() => setStep(2)} className="text-slate-400 hover:text-white text-sm underline">Edit Line Sub-Categories</button>
          <button onClick={() => setStep(3)} className="text-slate-400 hover:text-white text-sm underline">Edit Link Costs</button>
          <button onClick={() => setStep(5)} className="text-slate-400 hover:text-white text-sm underline">Edit Routes</button>
        </div>
        
        <div className="mb-6">
          <button onClick={() => setStep(5)} className="bg-slate-700 hover:bg-slate-600 text-white font-semibold py-2 px-6 rounded-lg">
            Back
          </button>
        </div>

        <div className="mb-6 flex flex-wrap gap-3">
          <button onClick={exportStatistics} className="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-6 rounded-lg">
            Export Statistics
          </button>
          <button onClick={exportEquationsPDF} className="bg-purple-600 hover:bg-purple-700 text-white font-semibold py-2 px-6 rounded-lg">
            Export Equations PDF
          </button>
          <button onClick={exportRouteSummary} className="bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-6 rounded-lg">
            Export Route Summary
          </button>
        </div>
        
        {routes.length > 0 && (
          <div className="mb-6 bg-slate-900 rounded-lg p-5 border border-slate-700">
            <div className="mb-3">
              <h3 className="text-white font-semibold">Route Cost Summary</h3>
            </div>
            <div className="space-y-4">
              {routes.map(route => {
                const routeLinks = links.filter(l => route.linkIds.includes(l.id));
                const totalLength = routeLinks.reduce((sum, l) => sum + (parseFloat(l.length) || 0), 0);
                
                // Calculate costs by category for this route (transmission only, first 7 categories)
                const routeCategoryTotals = {};
                let transmissionSubtotal = 0;
                
                categories.slice(0, 7).forEach(mc => {
                  if (subCategories[mc.key]) {
                    let categoryTotal = 0;
                    subCategories[mc.key].forEach((sc, si) => {
                      let subCatTotal = 0;
                      
                      // Special handling for Land Acquisition - use route-based value
                      if (mc.key === 'rightOfWay' && si === 0) {
                        subCatTotal = parseFloat(routeLandAcquisition[route.id]) || 0;
                      } else {
                        route.linkIds.forEach(linkId => {
                          const key = `${mc.key}_${si}_${linkId}`;
                          subCatTotal += round(inputs[key]);
                        });
                      }
                      
                      categoryTotal += subCatTotal;
                    });
                    routeCategoryTotals[mc.key] = categoryTotal;
                    transmissionSubtotal += categoryTotal;
                  }
                });
                
                // Calculate escalation and contingency factors
                const totalYears = escalationYears + (escalationMonths / 12);
                const escalationFactor = Math.pow(1 + (escalationRate / 100), totalYears);
                const contingencyFactor = 1 + (contingency / 100);
                
                // Apply escalation and contingency to each category, then round up to nearest 1,000
                const routeCategoryTotalsAdjusted = {};
                let totalEscalation = 0;
                let totalContingency = 0;
                
                categories.slice(0, 7).forEach(mc => {
                  const baseValue = routeCategoryTotals[mc.key] || 0;
                  
                  // Apply escalation
                  const afterEscalation = baseValue * escalationFactor;
                  const categoryEscalation = afterEscalation - baseValue;
                  totalEscalation += categoryEscalation;
                  
                  // Apply contingency
                  const afterContingency = afterEscalation * contingencyFactor;
                  const categoryContingency = afterContingency - afterEscalation;
                  totalContingency += categoryContingency;
                  
                  // Round final category value up to nearest 1,000
                  routeCategoryTotalsAdjusted[mc.key] = Math.ceil(afterContingency / 1000) * 1000;
                });
                
                const routeEscalation = totalEscalation;
                const routeContingency = totalContingency;
                
                // Calculate transmission total from rounded category values
                const transmissionTotal = Object.values(routeCategoryTotalsAdjusted).reduce((sum, val) => sum + val, 0);
                
                // Get substation cost from state (same for all routes) and round up
                const routeSubstationCost = Math.ceil((parseFloat(substationCost) || 0) / 1000) * 1000;
                const routeTotal = transmissionTotal + routeSubstationCost;
                
                return (
                  <div key={route.id} className="bg-slate-800 rounded-lg p-4">
                    <div className="flex justify-between items-center mb-3">
                      <div>
                        <span className="text-white font-semibold text-lg">{route.name || 'Unnamed Route'}</span>
                        <span className="text-slate-400 text-sm ml-3">
                          ({routeLinks.map(l => l.link).join('-')} • {totalLength.toFixed(2)} mi)
                        </span>
                        <div className="text-slate-400 text-xs mt-1">
                          {totalLength > 0 && `Avg. $/mi: $${(transmissionTotal / totalLength).toLocaleString('en-US', {minimumFractionDigits: 0, maximumFractionDigits: 0})}`}
                        </div>
                        <div className="text-slate-400 text-xs mt-1">
                          {(() => {
                            // Calculate structure counts for this route
                            const structureCounts = {};
                            [1, 2, 3, 4, 5].forEach(n => {
                              const typeName = structureNames[`type${n}`];
                              if (typeName && typeName !== 'N/A') {
                                const count = routeLinks.reduce((sum, l) => sum + (parseFloat(l[`structureType${n}`]) || 0), 0);
                                if (count > 0) {
                                  structureCounts[typeName] = count;
                                }
                              }
                            });
                            
                            // Format structure counts
                            const structureText = Object.entries(structureCounts)
                              .map(([name, count]) => `${name}: ${count}`)
                              .join(', ');
                            
                            return structureText ? `Str. Count: ${structureText}` : 'No structures';
                          })()}
                        </div>
                      </div>
                      <span className="text-white font-bold text-xl">${routeTotal.toLocaleString('en-US')}</span>
                    </div>
                    
                    <div className="mt-3 pt-3 border-t border-slate-700">
                      <div className="space-y-1">
                        {categories.slice(0, 7).map(mc => {
                          const catTotal = routeCategoryTotalsAdjusted[mc.key] || 0;
                          return (
                            <div key={mc.key} className="flex justify-between text-sm">
                              <span className="text-slate-400">{mc.label}</span>
                              <span className="text-slate-300 text-right">${catTotal.toLocaleString('en-US')}</span>
                            </div>
                          );
                        })}
                        
                        {/* Transmission Line Cost (sum of first 7 categories with escalation + contingency applied) */}
                        <div className="flex justify-between text-sm pt-2 border-t border-slate-600">
                          <span className="text-slate-300 font-semibold">Transmission Line Cost</span>
                          <span className="text-white font-semibold text-right">${transmissionTotal.toLocaleString('en-US')}</span>
                        </div>
                        
                        <div className="flex justify-between text-sm pt-2">
                          <span className="text-slate-400">Estimated Substation Facilities Cost</span>
                          <span className="text-slate-300 text-right">${routeSubstationCost.toLocaleString('en-US')}</span>
                        </div>
                        
                        {/* Total Project Cost (all categories) */}
                        <div className="flex justify-between text-sm pt-2 mt-2 border-t border-slate-600">
                          <span className="text-white font-bold">Total Project Cost</span>
                          <span className="text-white font-bold text-right">${routeTotal.toLocaleString('en-US')}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        
        {total > 0 && (
          <div className="bg-slate-900 rounded-xl p-6 border border-slate-700">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-xl font-bold text-white">Cost Summary</h2>
              <button onClick={exportExcel} className="bg-green-600 hover:bg-green-700 text-white font-semibold py-2 px-4 rounded-lg">Export Excel</button>
            </div>
            <div className="space-y-3">
              {categories.slice(0, 7).map(mc => {
                const t = catTotal(mc.key);
                return t > 0 ? <div key={mc.key} className="flex justify-between py-2 border-b border-slate-700"><span className="text-slate-300">{mc.label}</span><span className="text-white font-semibold">${t.toLocaleString('en-US')}</span></div> : null;
              })}
              <div className="flex justify-between py-3 pt-4 border-t-2 border-blue-500"><span className="text-white font-bold">Estimated Total Transmission Cost</span><span className="text-white font-bold text-2xl">${transTotal().toLocaleString('en-US')}</span></div>
              {catTotal('substationUtility') > 0 && <div className="flex justify-between py-2 border-b border-slate-700"><span className="text-slate-300">Estimated Substation Facilities Cost</span><span className="text-white font-semibold">${catTotal('substationUtility').toLocaleString('en-US')}</span></div>}
              <div className="flex justify-between py-3 border-t border-slate-700"><span className="text-white font-bold">Estimated Total Project Costs</span><span className="text-white font-bold text-2xl">${total.toLocaleString('en-US')}</span></div>
            </div>
          </div>
        )}
        
        <div className="mt-6">
          <button onClick={() => setStep(5)} className="bg-slate-700 hover:bg-slate-600 text-white font-semibold py-2 px-6 rounded-lg">
            ← Go Back
          </button>
        </div>
      </div>
    </div>
  );
}

