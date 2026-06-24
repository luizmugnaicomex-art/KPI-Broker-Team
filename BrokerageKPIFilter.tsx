import React, { useMemo, useState, useRef, useEffect } from 'react';
import { Shipment } from './types';

interface FilterProps {
  shipments: Shipment[];
  activeFilters: {
    analyst: string;
    month: string;
    year: number | 'All';
    cargo: string[];
    uniqueDi: string;
  };
  onFilterChange: (filterName: string, value: any) => void;
}

const MultiSelectDropdown = ({ options, selected, onChange, label }: { options: string[], selected: string[], onChange: (val: string[]) => void, label: string }) => {
    const [isOpen, setIsOpen] = useState(false);
    const containerRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
                setIsOpen(false);
            }
        };
        document.addEventListener('mousedown', handleClickOutside);
        return () => document.removeEventListener('mousedown', handleClickOutside);
    }, []);

    const toggleOption = (option: string) => {
        const newSelected = selected.includes(option)
            ? selected.filter(item => item !== option)
            : [...selected, option];
        onChange(newSelected);
    };

    return (
        <div className="filter-group" ref={containerRef} style={{ position: 'relative' }}>
            <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '0.25rem' }}>{label}</label>
            <div
                onClick={() => setIsOpen(!isOpen)}
                style={{
                    background: 'rgba(15, 23, 42, 0.96)',
                    border: '1px solid rgba(148, 163, 184, 0.5)',
                    padding: '0.4rem 0.6rem',
                    borderRadius: '8px',
                    color: '#fff',
                    cursor: 'pointer',
                    fontSize: '0.8rem',
                    minWidth: '200px',
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center'
                }}
            >
                <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '170px' }}>
                    {selected.length === 0 ? 'All' : selected.length === 1 ? selected[0] : `${selected.length} Selected`}
                </span>
                <span className="material-symbols-outlined" style={{ fontSize: '1.2rem' }}>arrow_drop_down</span>
            </div>
            {isOpen && (
                <div style={{
                    position: 'absolute',
                    top: '100%',
                    left: 0,
                    width: '100%',
                    maxHeight: '300px',
                    overflowY: 'auto',
                    background: '#0f172a',
                    border: '1px solid rgba(148, 163, 184, 0.5)',
                    borderRadius: '8px',
                    zIndex: 50,
                    marginTop: '4px',
                    boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.5)'
                }}>
                    <div
                        onClick={() => onChange([])}
                        style={{ padding: '8px 12px', cursor: 'pointer', fontSize: '0.8rem', borderBottom: '1px solid rgba(255,255,255,0.05)', color: selected.length === 0 ? '#60a5fa' : '#9ca3af', fontWeight: selected.length === 0 ? 'bold' : 'normal' }}
                    >
                        All
                    </div>
                    {options.map(option => (
                        <div
                            key={option}
                            onClick={() => toggleOption(option)}
                            style={{
                                padding: '8px 12px',
                                cursor: 'pointer',
                                fontSize: '0.8rem',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px',
                                background: selected.includes(option) ? 'rgba(59, 130, 246, 0.15)' : 'transparent',
                                color: selected.includes(option) ? '#fff' : '#cbd5e1'
                            }}
                        >
                            <span className="material-symbols-outlined" style={{ fontSize: '1rem', color: selected.includes(option) ? '#60a5fa' : 'transparent', border: selected.includes(option) ? 'none' : '1px solid #475569', borderRadius: '3px', width: '16px', height: '16px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                check
                            </span>
                            {option}
                        </div>
                    ))}
                </div>
            )}
        </div>
    );
};

const BrokerageKPIFilter: React.FC<FilterProps> = ({ shipments, activeFilters, onFilterChange }) => {
    const years = useMemo(() => {
        const yearSet = new Set(shipments.map(s => s.diRegistrationDate ? new Date(s.diRegistrationDate).getFullYear() : null).filter(y => y !== null));
        // Explicitly ensure 2026 is in the list
        yearSet.add(2026);
        const uniqueYears = Array.from(yearSet);
        return uniqueYears.sort((a, b) => (b as number) - (a as number));
    }, [shipments]);

    const analysts = useMemo(() => {
        const toTitleCase = (str: string | undefined): string => {
             if (!str) return '';
             return str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
        };
        return Array.from(new Set(shipments.map(s => toTitleCase(s.technicianResponsibleBrazil)).filter(Boolean))).sort();
    }, [shipments]);

    const cargos = useMemo(() => {
        return Array.from(new Set(shipments.map(s => s.typeOfCargo).filter(Boolean))).sort();
    }, [shipments]);

    const months = [
        { value: '01', label: 'Jan' }, { value: '02', label: 'Feb' }, { value: '03', label: 'Mar' },
        { value: '04', label: 'Apr' }, { value: '05', label: 'May' }, { value: '06', label: 'Jun' },
        { value: '07', label: 'Jul' }, { value: '08', label: 'Aug' }, { value: '09', label: 'Sep' },
        { value: '10', label: 'Oct' }, { value: '11', label: 'Nov' }, { value: '12', label: 'Dec' }
    ];

    const selectStyle = {
        background: 'rgba(15, 23, 42, 0.96)',
        border: '1px solid rgba(148, 163, 184, 0.5)',
        padding: '0.4rem',
        borderRadius: '8px',
        color: '#fff',
        fontSize: '0.8rem',
        height: '32px',
        minWidth: '80px'
    };

    return (
        <div className="brokerage-filters" style={{ display: 'flex', flexWrap: 'wrap', gap: '1rem', paddingBottom: '0.5rem', alignItems: 'flex-end' }}>
            <div className="filter-group">
                <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '0.25rem' }}>Year</label>
                <select value={activeFilters.year} onChange={(e) => onFilterChange('year', e.target.value === 'All' ? 'All' : Number(e.target.value))} style={selectStyle}>
                    <option value="All">All</option>
                    {years.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
            </div>
             <div className="filter-group">
                <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '0.25rem' }}>Month</label>
                <select value={activeFilters.month} onChange={(e) => onFilterChange('month', e.target.value)} style={selectStyle}>
                    <option value="All">All</option>
                    {months.map(m => <option key={m.value} value={m.value}>{m.label}</option>)}
                </select>
            </div>
            <div className="filter-group">
                <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '0.25rem' }}>Analyst</label>
                <select value={activeFilters.analyst} onChange={(e) => onFilterChange('analyst', e.target.value)} style={{...selectStyle, minWidth: '150px'}}>
                    <option value="All">All</option>
                    {analysts.map(a => <option key={a} value={a}>{a}</option>)}
                </select>
            </div>

            <MultiSelectDropdown
                label="Cargo"
                options={cargos as string[]}
                selected={activeFilters.cargo}
                onChange={(val) => onFilterChange('cargo', val)}
            />

            <div className="filter-group">
                <label style={{ display: 'block', fontSize: '0.75rem', color: '#94a3b8', marginBottom: '0.25rem' }}>Unique DI</label>
                <select value={activeFilters.uniqueDi} onChange={(e) => onFilterChange('uniqueDi', e.target.value)} style={selectStyle}>
                    <option value="All">All</option>
                    <option value="Yes">Yes</option>
                    <option value="No">No</option>
                </select>
            </div>
        </div>
    );
};

export default BrokerageKPIFilter;