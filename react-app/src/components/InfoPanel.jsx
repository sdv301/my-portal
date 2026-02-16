// src/components/InfoPanel.jsx
import React, { useState, useEffect, useCallback } from 'react';
import './InfoPanel.css';

const DEFAULT_FILTERS = { startDate: '', endDate: '' };

const InfoPanel = ({ district, filters = {}, isFiltersExpanded, onPanelToggle }) => {
  const [districtInfo, setDistrictInfo] = useState(null);
  const [districtData, setDistrictData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('inventory');
  const [isExpanded, setIsExpanded] = useState(false);

  useEffect(() => {
    if (district?.id) {
      setIsExpanded(true);
      if (onPanelToggle) onPanelToggle(true);
    }
  }, [district?.id]);

  useEffect(() => {
    if (isFiltersExpanded && isExpanded) togglePanel();
  }, [isFiltersExpanded]);

  useEffect(() => {
    if (!district?.id) {
      setDistrictInfo(null); setDistrictData(null); return;
    }

    const loadData = async () => {
      setLoading(true);
      try {
        const encodedId = encodeURIComponent(district.id);
        const start = filters.startDate || '';
        const end = filters.endDate || '';

        // ДОБАВЛЕНО: Передаем даты с фильтра в запрос
        const [infoRes, dataRes] = await Promise.allSettled([
          fetch(`http://localhost:5000/api/district/${encodedId}`),
          fetch(`http://localhost:5000/api/district/${encodedId}/data?startDate=${start}&endDate=${end}`)
        ]);

        if (infoRes.status === 'fulfilled' && infoRes.value.ok) setDistrictInfo(await infoRes.value.json());
        if (dataRes.status === 'fulfilled' && dataRes.value.ok) setDistrictData(await dataRes.value.json());
      } catch (error) {
        console.error('Ошибка:', error);
      } finally {
        setLoading(false);
      }
    };
    loadData();
  }, [district, filters]); // Панель будет обновляться при смене дат в фильтре!

  const togglePanel = useCallback(() => {
    const newState = !isExpanded;
    setIsExpanded(newState);
    if (onPanelToggle) onPanelToggle(newState);
  }, [isExpanded, onPanelToggle]);

  const formatCurrency = (value) => {
    return new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB' }).format(value);
  };

  return (
    <div className={`info-panel ${isExpanded ? 'expanded' : 'collapsed'}`}>
      <button className="panel-toggle-btn" onClick={togglePanel}>
        <span className="toggle-icon">{isExpanded ? '▶' : '◀'}</span>
      </button>

      {!district?.id ? (
        <div className="empty-state">
          <h3>👆 Выберите район</h3>
          <p>Кликните на карту для информации</p>
        </div>
      ) : loading ? (
        <div className="loading-state">
          <h3 style={{color: 'var(--primary-color)'}}>Загрузка...</h3>
        </div>
      ) : (
        <>
          <div className="panel-header">
            <h2>{districtInfo?.name || district.name}</h2>
            {districtData?.statistics?.total_cost > 0 && (
              <div className="total-cost-badge">
                <small>Всего выделено МЦ:</small>
                <strong>{formatCurrency(districtData.statistics.total_cost)}</strong>
                
                {/* ДОБАВЛЕНО: Отображение дат из базы */}
                {districtData.statistics.earliest_date && (
                  <div style={{fontSize: '11px', color: 'var(--text-secondary)', marginTop: '5px', fontWeight: '500'}}>
                    📅 За период: {districtData.statistics.earliest_date}
                    {districtData.statistics.earliest_date !== districtData.statistics.latest_date ? ` — ${districtData.statistics.latest_date}` : ''} гг.
                  </div>
                )}
              </div>
            )}
          </div>

          <div className="tabs">
            {['inventory', 'overview'].map(tab => (
              <button key={tab} className={activeTab === tab ? 'active' : ''} onClick={() => setActiveTab(tab)}>
                {tab === 'inventory' ? 'Резервы (МЦ)' : 'Обзор района'}
              </button>
            ))}
          </div>

          <div className="tab-content">
            {activeTab === 'inventory' && (
              <div className="inventory-list">
                {districtData?.inventory && Object.keys(districtData.inventory).length > 0 ? (
                  Object.entries(districtData.inventory).map(([category, items]) => (
                    <div key={category} className="inventory-category-block">
                      <h4 className="category-title">{category}</h4>
                      <div className="table-responsive">
                        <table className="inventory-table">
                          <thead>
                            <tr>
                              <th>Наименование</th>
                              <th style={{textAlign: 'center'}}>Кол-во</th>
                              <th style={{textAlign: 'right'}}>Стоимость</th>
                            </tr>
                          </thead>
                          <tbody>
                            {items.map((item, idx) => (
                              <tr key={idx}>
                                <td>
                                  <div className="item-name">{item.name}</div>
                                  <div className="item-year">Выдано: {item.year} г.</div>
                                </td>
                                <td style={{textAlign: 'center', whiteSpace: 'nowrap'}}>
                                  <span className="qty-badge">{item.quantity} {item.unit}</span>
                                </td>
                                <td style={{textAlign: 'right', whiteSpace: 'nowrap', fontWeight: '500'}}>
                                  {formatCurrency(item.cost)}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ))
                ) : (
                  <div className="empty-state" style={{marginTop: '40px'}}>
                    <p style={{fontSize: '32px', margin: '0 0 10px 0'}}>📦</p>
                    <p>Нет выданного имущества за выбранный период</p>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'overview' && (
              <div className="stats-grid">
                <div className="stat-card">
                  <div className="stat-label">Население</div>
                  <div className="stat-value">{districtInfo?.population?.toLocaleString() || 'Н/Д'}</div>
                </div>
                <div className="stat-card">
                  <div className="stat-label">Площадь</div>
                  <div className="stat-value">{districtInfo?.area_km2?.toLocaleString() || 'Н/Д'}</div>
                  <div className="stat-unit">км²</div>
                </div>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
};

export default InfoPanel;