export type DIChannel = 'Green' | 'Yellow' | 'Red' | 'Gray' | 'Pending' | 'Other';

export interface Shipment {
  diRegistrationDate?: string | Date;
  technicianResponsibleBrazil?: string;
  typeOfCargo?: string;
  invoiceValue?: number;
  cargoPresenceDate?: string | Date;
  greenChannelOrDeliveryAuthorizedDate?: string | Date;
  shipmentType?: string;
  actualEtd?: string | Date;
  actualEta?: string | Date;
  incoterm?: string;
  channel?: DIChannel;
  diNumber?: string;
  [key: string]: any;
}