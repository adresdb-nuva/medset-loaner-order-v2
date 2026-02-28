export interface MedicalSet {
  id: string;
  name: string;
  code?: string;
  description?: string;
  region: string;
}

export interface RegionData {
  name: string;
  sets: MedicalSet[];
}

export interface OrderItem {
  setId: string;
  setName: string;
  setCode?: string;
  region: string;
}

export interface OrderMetadata {
  hospital: string;
  date: string;
  surgeon: string;
  agentName: string;
  remarks: string;
  infoEmails: string;
}

export interface OrderHistoryItem {
  id: string;
  timestamp: string;
  items: OrderItem[];
  metadata: OrderMetadata;
}
