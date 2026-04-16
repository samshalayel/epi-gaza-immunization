const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

const ANTIGENS = [
  {key:'BCG',    label:'BCG'},
  {key:'OPV0',   label:'OPV 0 (Birth)'},
  {key:'Penta1', label:'Penta 1'},
  {key:'OPV1',   label:'OPV 1'},
  {key:'PCV1',   label:'PCV 1'},
  {key:'Penta2', label:'Penta 2'},
  {key:'OPV2',   label:'OPV 2'},
  {key:'PCV2',   label:'PCV 2'},
  {key:'Penta3', label:'Penta 3'},
  {key:'OPV3',   label:'OPV 3'},
  {key:'PCV3',   label:'PCV 3'},
  {key:'IPV',    label:'IPV'},
  {key:'MR1',    label:'MR 1 (MMR1)'},
  {key:'MR2',    label:'MR 2 (MMR2)'},
  {key:'VitA1',  label:'Vit A (1st)'},
  {key:'VitA2',  label:'Vit A (2nd)'},
];

const DROPOUTS = [
  {label:'Dropout 1: Penta1 → Penta3', num:'Penta1', den:'Penta3'},
  {label:'Dropout 2: BCG → MR1',        num:'BCG',    den:'MR1'},
  {label:'Dropout 3: MR1 → MR2',        num:'MR1',    den:'MR2'},
];
