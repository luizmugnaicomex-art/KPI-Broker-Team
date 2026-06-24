import { Shipment, DIChannel } from './types';

export const calculateDaysBetween = (d1: string | Date | undefined, d2: string | Date | undefined): number | null => {
  if (!d1 || !d2) return null;
  const date1 = new Date(d1);
  const date2 = new Date(d2);
  if (isNaN(date1.getTime()) || isNaN(date2.getTime())) return null;
  const diffTime = Math.abs(date2.getTime() - date1.getTime());
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
};

export const buildUniqueDIIndex = (shipments: Shipment[]) => {
  const map = new Map<string, { channel: DIChannel, shipments: Shipment[] }>();
  
  shipments.forEach((s) => {
    // Use diNumber for grouping, fallback if missing to ensure all items are counted
    const key = s.diNumber || `NO_DI_${Math.random()}`;
    
    if (!map.has(key)) {
      // Determine channel for the DI based on the shipment data
      // If channel is missing, default to 'Other' to avoid crashes
      const channel = (s.channel as DIChannel) || 'Other';
      map.set(key, { channel, shipments: [] });
    }
    
    map.get(key)!.shipments.push(s);
  });
  
  return map;
};