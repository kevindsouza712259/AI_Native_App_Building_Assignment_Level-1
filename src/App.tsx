/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import { FileSpreadsheet, ArrowRight, CheckCircle2, LayoutDashboard, BarChart3, Store, AlertCircle, CheckCircle, Sparkles, Download, FileText, Loader2, DollarSign, Target, ShoppingCart, Users, Percent, ArrowUp, ArrowDown, PackageX, Star } from 'lucide-react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, LineChart, Line, Legend } from 'recharts';
import { MultiSelect } from './components/MultiSelect';

export default function App() {
  const [salesFile, setSalesFile] = useState<File | null>(null);
  const [storeFile, setStoreFile] = useState<File | null>(null);
  const [isLoaded, setIsLoaded] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [successMsg, setSuccessMsg] = useState<string | null>(null);

  const [rawSalesData, setRawSalesData] = useState<any[] | null>(null);
  const [storeMap, setStoreMap] = useState<Map<string, any> | null>(null);
  const [filters, setFilters] = useState<{
    week_start_date: string[];
    region: string[];
    store_name: string[];
    city: string[];
    store_format: string[];
    product_category: string[];
  }>({
    week_start_date: [],
    region: [],
    store_name: [],
    city: [],
    store_format: [],
    product_category: [],
  });

  const [pendingFilters, setPendingFilters] = useState<{
    week_start_date: string[];
    region: string[];
    store_name: string[];
    city: string[];
    store_format: string[];
    product_category: string[];
  }>({
    week_start_date: [],
    region: [],
    store_name: [],
    city: [],
    store_format: [],
    product_category: [],
  });

  const [isMobileFiltersOpen, setIsMobileFiltersOpen] = useState(false);

  const handleApplyFilters = () => {
    setFilters(pendingFilters);
  };

  const handleResetFilters = () => {
    const empty = { week_start_date: [], region: [], store_name: [], city: [], store_format: [], product_category: [] };
    setPendingFilters(empty);
    setFilters(empty);
  };

  const handleSalesUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setSalesFile(e.target.files[0]);
      setIsLoaded(false); // Reset dashboard on new file upload
      setSuccessMsg(null);
    }
  };

  const handleStoreUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setStoreFile(e.target.files[0]);
      setIsLoaded(false); // Reset dashboard on new file upload
      setSuccessMsg(null);
    }
  };

  const handleLoadDashboard = async () => {
    if (!salesFile || !storeFile) {
      setErrorMsg("Please upload both required files.");
      return;
    }
    setErrorMsg(null);
    setSuccessMsg(null);
    setIsLoading(true);
    
    // Give React a tick to render the loading spinner
    await new Promise(resolve => setTimeout(resolve, 50));

    try {
      const salesArrayBuffer = await salesFile.arrayBuffer();
      const storeArrayBuffer = await storeFile.arrayBuffer();

      const salesWorkbook = XLSX.read(salesArrayBuffer, { type: 'array' });
      const storeWorkbook = XLSX.read(storeArrayBuffer, { type: 'array' });

      const storeSheetName = storeWorkbook.SheetNames[0];

      const cleanDataSheetName = 'Clean_Data';
      if (!salesWorkbook.Sheets[cleanDataSheetName]) {
        throw new Error("retail_weekly_sales.xlsx must contain a sheet named 'Clean_Data'.");
      }

      const salesData = XLSX.utils.sheet_to_json(salesWorkbook.Sheets[cleanDataSheetName]);
      const storeData = XLSX.utils.sheet_to_json(storeWorkbook.Sheets[storeSheetName]);

      if (!salesData.length || !storeData.length) {
        throw new Error("One or both uploaded files are empty.");
      }

      let totalNetSales = 0;
      let totalSalesTarget = 0;
      let totalReturnsAmount = 0;
      let totalDiscountAmount = 0;
      let totalGrossSales = 0;
      let totalTransactions = 0;
      let totalFootfall = 0;
      let sumCustomerRating = 0;
      let countCustomerRating = 0;

      const storeMapData = new Map();
      storeData.forEach((store: any) => {
        if (store.store_id !== undefined) {
          storeMapData.set(String(store.store_id), store);
        }
      });

      setRawSalesData(salesData);
      setStoreMap(storeMapData);

      let loadedMsg = "Data loaded successfully!";
      if (salesData.length === 1919) {
          loadedMsg = `Successfully loaded and verified 1919 records from retail_weekly_sales.xlsx.`;
      } else {
          loadedMsg = `Loaded data. Note: Expected 1919 records, but found ${salesData.length}.`;
      }
      setSuccessMsg(loadedMsg);
      setIsLoaded(true);
    } catch (err: any) {
      setErrorMsg(err.message || "Failed to parse the uploaded files. Ensure they are valid Excel files.");
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
  };

  const formatNumber = (value: number) => {
    return new Intl.NumberFormat('en-US', { maximumFractionDigits: 0 }).format(value);
  };

  const formatPercent = (value: number) => {
    return new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value) + '%';
  };

  const formatDecimal = (value: number) => {
    return new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
  };

  const dashboardData = React.useMemo(() => {
    if (!rawSalesData || !storeMap) return null;

    let totalNetSales = 0;
    let totalSalesTarget = 0;
    let totalReturnsAmount = 0;
    let totalDiscountAmount = 0;
    let totalGrossSales = 0;
    let totalTransactions = 0;
    let totalFootfall = 0;
    let sumCustomerRating = 0;
    let countCustomerRating = 0;
    let totalStockouts = 0;
    let countStockouts = 0;

    const revenueByDateMap = new Map();
    const revenueByStoreMap = new Map();
    const revenueByRegionMap = new Map();
    const revenueByCategoryMap = new Map();
    const stockoutsByStoreMap = new Map();

    const regionsSet = new Set<string>();
    const datesSet = new Set<string>();
    const storesSet = new Set<string>();
    const citiesSet = new Set<string>();
    const formatsSet = new Set<string>();
    const categoriesSet = new Set<string>();

    let globalTotalStockouts = 0;

    // First pass: extract all available filter values from raw data
    rawSalesData.forEach((sale: any) => {
      const storeId = String(sale.store_id || '');
      const storeInfo = storeMap.get(storeId) || {};
      
      globalTotalStockouts += parseFloat(sale.stockouts) || 0;
      
      let dateStr = sale.week_start_date || 'Unknown';
      if (typeof dateStr === 'number') {
         const d = new Date(Math.round((dateStr - 25569)*86400*1000));
         dateStr = d.toISOString().split('T')[0];
      } else {
         dateStr = String(dateStr);
      }

      const region = String(sale.region || 'Unknown');
      const storeName = String(storeInfo.store_name || `Store ${storeId}`);
      const city = String(storeInfo.city || 'Unknown');
      const storeFormat = String(storeInfo.store_format || 'Unknown');
      const category = String(sale.product_category || 'Unknown');

      datesSet.add(dateStr);
      regionsSet.add(region);
      storesSet.add(storeName);
      citiesSet.add(city);
      formatsSet.add(storeFormat);
      categoriesSet.add(category);
    });

    // Apply all filters
    const filteredData = rawSalesData.filter((sale: any) => {
      const storeId = String(sale.store_id || '');
      const storeInfo = storeMap.get(storeId) || {};
      
      let dateStr = sale.week_start_date || 'Unknown';
      if (typeof dateStr === 'number') {
         const d = new Date(Math.round((dateStr - 25569)*86400*1000));
         dateStr = d.toISOString().split('T')[0];
      } else {
         dateStr = String(dateStr);
      }

      const region = String(sale.region || 'Unknown');
      const storeName = String(storeInfo.store_name || `Store ${storeId}`);
      const city = String(storeInfo.city || 'Unknown');
      const storeFormat = String(storeInfo.store_format || 'Unknown');
      const category = String(sale.product_category || 'Unknown');

      if (filters.week_start_date.length > 0 && !filters.week_start_date.includes(dateStr)) return false;
      if (filters.region.length > 0 && !filters.region.includes(region)) return false;
      if (filters.store_name.length > 0 && !filters.store_name.includes(storeName)) return false;
      if (filters.city.length > 0 && !filters.city.includes(city)) return false;
      if (filters.store_format.length > 0 && !filters.store_format.includes(storeFormat)) return false;
      if (filters.product_category.length > 0 && !filters.product_category.includes(category)) return false;
      
      return true;
    });

    filteredData.forEach((sale: any) => {
      const storeId = String(sale.store_id || '');
      const netSales = parseFloat(sale.net_sales) || 0;
      const salesTarget = parseFloat(sale.sales_target) || 0;
      const returnsAmount = parseFloat(sale.returns_amount) || 0;
      const discountAmount = parseFloat(sale.discount_amount) || 0;
      const grossSales = parseFloat(sale.gross_sales) || 0;
      const transactions = parseFloat(sale.transactions) || 0;
      const footfall = parseFloat(sale.footfall) || 0;
      const customerRating = parseFloat(sale.customer_rating);
      const region = String(sale.region || 'Unknown');
      const category = String(sale.product_category || 'Unknown');
      const stockouts = parseFloat(sale.stockouts) || 0;
      
      let dateStr = sale.week_start_date || 'Unknown';
      if (typeof dateStr === 'number') {
         const d = new Date(Math.round((dateStr - 25569)*86400*1000));
         dateStr = d.toISOString().split('T')[0];
      } else {
         dateStr = String(dateStr);
      }

      totalNetSales += netSales;
      totalSalesTarget += salesTarget;
      totalReturnsAmount += returnsAmount;
      totalDiscountAmount += discountAmount;
      totalGrossSales += grossSales;
      totalTransactions += transactions;
      totalFootfall += footfall;
      totalStockouts += stockouts;
      countStockouts++;
      
      if (!isNaN(customerRating)) {
         sumCustomerRating += customerRating;
         countCustomerRating++;
      }

      if (!revenueByDateMap.has(dateStr)) {
        revenueByDateMap.set(dateStr, { date: dateStr, revenue: 0, target: 0 });
      }
      revenueByDateMap.get(dateStr).revenue += netSales;
      revenueByDateMap.get(dateStr).target += salesTarget;

      if (!revenueByRegionMap.has(region)) {
        revenueByRegionMap.set(region, { name: region, revenue: 0, stockouts: 0 });
      }
      revenueByRegionMap.get(region).revenue += netSales;
      revenueByRegionMap.get(region).stockouts += stockouts;

      if (!revenueByCategoryMap.has(category)) {
        revenueByCategoryMap.set(category, { name: category, revenue: 0, returns: 0 });
      }
      revenueByCategoryMap.get(category).revenue += netSales;
      revenueByCategoryMap.get(category).returns += returnsAmount;

      const storeInfo = storeMap.get(storeId) || {};
      const storeName = storeInfo.store_name || `Store ${storeId}`;

      if (!revenueByStoreMap.has(storeName)) {
        revenueByStoreMap.set(storeName, { name: storeName, revenue: 0, target: 0 });
      }
      revenueByStoreMap.get(storeName).revenue += netSales;
      revenueByStoreMap.get(storeName).target += salesTarget;

      if (!stockoutsByStoreMap.has(storeName)) {
        stockoutsByStoreMap.set(storeName, { name: storeName, stockouts: 0 });
      }
      stockoutsByStoreMap.get(storeName).stockouts += stockouts;
    });

    const targetAchievement = totalSalesTarget > 0 ? (totalNetSales / totalSalesTarget) * 100 : 0;
    const avgTransactionValue = totalTransactions > 0 ? totalNetSales / totalTransactions : 0;
    const returnRate = totalNetSales > 0 ? (totalReturnsAmount / totalNetSales) * 100 : 0;
    const discountRate = totalGrossSales > 0 ? (totalDiscountAmount / totalGrossSales) * 100 : 0;
    const avgCustomerRating = countCustomerRating > 0 ? sumCustomerRating / countCustomerRating : 0;

    const chartDataRegion = Array.from(revenueByRegionMap.values()).sort((a, b) => b.revenue - a.revenue);
    const chartDataStoreAll = Array.from(revenueByStoreMap.values()).sort((a, b) => b.revenue - a.revenue);
    const chartDataCategory = Array.from(revenueByCategoryMap.values()).sort((a, b) => b.revenue - a.revenue);

    // Business Insights Calculation
    const insights: string[] = [];

    if (chartDataRegion.length > 0) {
      const bestRegion = chartDataRegion[0].name;
      insights.push(`${bestRegion} region generated the highest revenue.`);
      
      const worstRegion = chartDataRegion[chartDataRegion.length - 1].name;
      if (bestRegion !== worstRegion) {
         insights.push(`${worstRegion} region generated the lowest revenue.`);
      }
    }

    if (chartDataStoreAll.length > 0) {
      const bestStore = chartDataStoreAll[0].name;
      insights.push(`${bestStore} is the best performing store.`);
      
      const storesMissingTarget = chartDataStoreAll.filter(s => s.revenue < s.target);
      if (storesMissingTarget.length > 0) {
        if (storesMissingTarget.length <= 3) {
          insights.push(`Stores missing sales target: ${storesMissingTarget.map(s => s.name).join(', ')}.`);
        } else {
          insights.push(`${storesMissingTarget.length} stores are currently below their sales target.`);
        }
      } else {
        insights.push(`All stores are meeting or exceeding their sales targets.`);
      }
    }

    if (chartDataCategory.length > 0) {
      let maxReturnRate = -1;
      let worstCategory = '';
      chartDataCategory.forEach(c => {
         const catReturnRate = c.revenue > 0 ? c.returns / c.revenue : 0;
         if (catReturnRate > maxReturnRate) {
            maxReturnRate = catReturnRate;
            worstCategory = c.name;
         }
      });
      if (worstCategory) {
        insights.push(`${worstCategory} has the highest return rate.`);
      }
    }

    const regionsByStockouts = Array.from(revenueByRegionMap.values()).sort((a, b) => b.stockouts - a.stockouts);
    if (regionsByStockouts.length > 0 && regionsByStockouts[0].stockouts > 0) {
      insights.push(`${regionsByStockouts[0].name} region experienced the highest stockouts.`);
    }

    if (avgCustomerRating > 0) {
      insights.push(`Average customer rating is ${avgCustomerRating.toFixed(1)}/5.`);
    }

    insights.push(`Overall Target Achievement is ${targetAchievement.toFixed(1)}%.`);

    const avgStockouts = countStockouts > 0 ? totalStockouts / countStockouts : 0;
    const globalAvgStockouts = rawSalesData.length > 0 ? globalTotalStockouts / rawSalesData.length : 0;

    return {
      filteredData,
      filteredRecordCount: filteredData.length,
      totalNetSales, targetAchievement, avgTransactionValue, returnRate, discountRate, totalTransactions, totalFootfall, avgCustomerRating,
      totalStockouts, avgStockouts, globalAvgStockouts,
      chartDataDate: Array.from(revenueByDateMap.values()).sort((a, b) => a.date.localeCompare(b.date)),
      chartDataStore: chartDataStoreAll.slice(0, 10),
      chartDataRegion,
      chartDataCategory,
      chartDataStockouts: Array.from(stockoutsByStoreMap.values()).sort((a, b) => b.stockouts - a.stockouts).slice(0, 10),
      filterOptions: {
        week_start_date: Array.from(datesSet).sort(),
        region: Array.from(regionsSet).sort(),
        store_name: Array.from(storesSet).sort(),
        city: Array.from(citiesSet).sort(),
        store_format: Array.from(formatsSet).sort(),
        product_category: Array.from(categoriesSet).sort()
      },
      insights
    };
  }, [rawSalesData, storeMap, filters]);

  const handleExportCSV = () => {
    if (!dashboardData?.filteredData) return;
    const worksheet = XLSX.utils.json_to_sheet(dashboardData.filteredData);
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'filtered_sales_data.csv');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleDownloadInsights = () => {
    if (!dashboardData?.insights) return;
    const textContent = "Business Insight Summary\n========================\n\n" + dashboardData.insights.map((i: string) => `• ${i}`).join('\n');
    const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'business_insights.txt');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (isLoaded && dashboardData) {
    return (
      <div className="min-h-screen bg-slate-950 text-slate-100 font-sans relative flex flex-col overflow-x-hidden">
        {/* Background Mesh Orbs */}
        <div className="fixed top-[-10%] left-[-10%] w-[500px] h-[500px] bg-blue-600/20 rounded-full blur-[120px] pointer-events-none z-0"></div>
        <div className="fixed bottom-[-10%] right-[-10%] w-[600px] h-[600px] bg-indigo-600/20 rounded-full blur-[150px] pointer-events-none z-0"></div>

        {/* Dashboard Header */}
        <header className="h-16 border-b border-white/10 flex items-center justify-between px-6 backdrop-blur-md bg-white/5 z-10">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-indigo-500 rounded-lg flex items-center justify-center">
              <LayoutDashboard className="w-5 h-5 text-white" />
            </div>
            <h1 className="text-lg font-bold text-white tracking-tight">
              RETAIL<span className="text-indigo-400">IQ</span>
            </h1>
          </div>
          
          <div className="flex items-center gap-4">
            <button 
              onClick={() => { setIsLoaded(false); setSuccessMsg(null); }}
              className="text-sm text-slate-400 hover:text-white font-medium px-4 py-2 rounded-lg bg-white/5 hover:bg-white/10 border border-white/10 transition-colors"
            >
              Upload New Data
            </button>
          </div>
        </header>        {/* Dashboard Content Structure */}
        <main className="flex-1 p-4 md:p-6 lg:p-8 w-full max-w-[100vw] mx-auto space-y-6 z-10">
          
          {successMsg && (
            <div className="bg-emerald-500/10 border border-emerald-500/50 text-emerald-400 px-4 py-3 rounded-xl flex items-center gap-3">
              <CheckCircle className="w-5 h-5" />
              <p className="font-medium">{successMsg}</p>
            </div>
          )}

          {/* Top: KPI Cards */}
          <div className="grid grid-cols-2 md:grid-cols-4 xl:grid-cols-8 gap-4">
            {/* KPI 1: Net Sales */}
            <div className="group relative" title="Total net sales after returns and discounts">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2">
                   <DollarSign className="w-4 h-4 text-indigo-400" />
                   <p className="text-xs font-medium text-slate-400 truncate">Total Net Sales</p>
                 </div>
                 <p className="text-lg font-semibold text-white truncate">{formatCurrency(dashboardData.totalNetSales)}</p>
               </div>
            </div>
            {/* KPI 2: Target Achieved */}
            <div className="group relative" title="Percentage of net sales compared to sales target">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2 justify-between">
                   <div className="flex items-center gap-2">
                     <Target className="w-4 h-4 text-indigo-400" />
                     <p className="text-xs font-medium text-slate-400 truncate">Target Achieved</p>
                   </div>
                   {dashboardData.targetAchievement >= 100 ? (
                     <ArrowUp className="w-3 h-3 text-emerald-400" />
                   ) : (
                     <ArrowDown className="w-3 h-3 text-rose-400" />
                   )}
                 </div>
                 <p className={`text-lg font-semibold truncate ${dashboardData.targetAchievement >= 100 ? 'text-emerald-400' : 'text-rose-400'}`}>
                   {formatPercent(dashboardData.targetAchievement)}
                 </p>
                 <div className="w-full bg-white/10 rounded-full h-1.5 mt-auto">
                   <div className={`h-1.5 rounded-full ${dashboardData.targetAchievement >= 100 ? 'bg-emerald-400' : 'bg-indigo-500'}`} style={{ width: `${Math.min(100, dashboardData.targetAchievement)}%` }}></div>
                 </div>
               </div>
            </div>
            {/* KPI 3: Avg Order Value */}
            <div className="group relative" title="Average revenue per transaction">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2">
                   <ShoppingCart className="w-4 h-4 text-indigo-400" />
                   <p className="text-xs font-medium text-slate-400 truncate">Avg Order Value</p>
                 </div>
                 <p className="text-lg font-semibold text-white truncate">{formatCurrency(dashboardData.avgTransactionValue)}</p>
               </div>
            </div>
            {/* KPI 4: Return Rate */}
            <div className="group relative" title="Percentage of sales returned. Goal is < 5%">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2 justify-between">
                   <div className="flex items-center gap-2">
                     <PackageX className="w-4 h-4 text-indigo-400" />
                     <p className="text-xs font-medium text-slate-400 truncate">Return Rate</p>
                   </div>
                   {dashboardData.returnRate > 5 ? (
                     <ArrowUp className="w-3 h-3 text-rose-400" />
                   ) : (
                     <ArrowDown className="w-3 h-3 text-emerald-400" />
                   )}
                 </div>
                 <p className={`text-lg font-semibold truncate ${dashboardData.returnRate > 5 ? 'text-rose-400' : 'text-emerald-400'}`}>
                   {formatPercent(dashboardData.returnRate)}
                 </p>
               </div>
            </div>
            {/* KPI 5: Discount Rate */}
            <div className="group relative" title="Percentage of gross sales given as discounts">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2">
                   <Percent className="w-4 h-4 text-indigo-400" />
                   <p className="text-xs font-medium text-slate-400 truncate">Discount Rate</p>
                 </div>
                 <p className="text-lg font-semibold text-white truncate">{formatPercent(dashboardData.discountRate)}</p>
               </div>
            </div>
            {/* KPI 6: Transactions */}
            <div className="group relative" title="Total number of customer transactions">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2">
                   <LayoutDashboard className="w-4 h-4 text-indigo-400" />
                   <p className="text-xs font-medium text-slate-400 truncate">Transactions</p>
                 </div>
                 <p className="text-lg font-semibold text-white truncate">{formatNumber(dashboardData.totalTransactions)}</p>
               </div>
            </div>
            {/* KPI 7: Stockouts */}
            <div className="group relative" title="Average stockout risk per store">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2 justify-between">
                   <div className="flex items-center gap-2">
                     <AlertCircle className="w-4 h-4 text-indigo-400" />
                     <p className="text-xs font-medium text-slate-400 truncate">Stockout Risk</p>
                   </div>
                   {dashboardData.avgStockouts > dashboardData.globalAvgStockouts ? (
                     <ArrowUp className="w-3 h-3 text-rose-400" title="Higher than global average" />
                   ) : (
                     <ArrowDown className="w-3 h-3 text-emerald-400" title="Lower than global average" />
                   )}
                 </div>
                 <p className={`text-lg font-semibold truncate ${dashboardData.avgStockouts > dashboardData.globalAvgStockouts ? 'text-rose-400' : 'text-emerald-400'}`}>
                   {formatDecimal(dashboardData.avgStockouts)} avg
                 </p>
               </div>
            </div>
            {/* KPI 8: Avg Rating */}
            <div className="group relative" title="Average customer satisfaction rating (out of 5)">
               <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
               <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-4 rounded-2xl flex flex-col gap-2 h-full">
                 <div className="flex items-center gap-2">
                   <Star className="w-4 h-4 text-amber-400" />
                   <p className="text-xs font-medium text-slate-400 truncate">Avg Rating</p>
                 </div>
                 <p className="text-lg font-semibold text-white truncate">{formatDecimal(dashboardData.avgCustomerRating)} / 5</p>
               </div>
            </div>
          </div>

          <div className="flex flex-col gap-6 w-full">
            
            {/* Filters Row */}
            <section className="w-full flex-shrink-0 flex flex-col gap-4 z-20">
              <div className="bg-white/5 backdrop-blur-xl border border-white/10 p-5 rounded-2xl flex flex-col gap-4">
                <div 
                  className="flex items-center justify-between cursor-pointer lg:cursor-default"
                  onClick={() => setIsMobileFiltersOpen(!isMobileFiltersOpen)}
                >
                  <h3 className="text-sm font-semibold text-white uppercase tracking-wider">Filters</h3>
                  <button className="lg:hidden text-slate-400">
                    {isMobileFiltersOpen ? 'Hide' : 'Show'}
                  </button>
                </div>
                
                <div className={`flex-col lg:flex-row flex-wrap items-start lg:items-end gap-4 ${isMobileFiltersOpen ? 'flex' : 'hidden'} lg:flex`}>
                  <div className="grid grid-cols-1 md:grid-cols-3 xl:grid-cols-6 gap-4 flex-1 w-full">
                    <MultiSelect
                      label="Week Start Date"
                      options={dashboardData.filterOptions.week_start_date}
                      selected={pendingFilters.week_start_date}
                      onChange={(selected) => {
                        setPendingFilters(f => ({ ...f, week_start_date: selected }));
                        setFilters(f => ({ ...f, week_start_date: selected }));
                      }}
                    />
                    <MultiSelect
                      label="Region"
                      options={dashboardData.filterOptions.region}
                      selected={pendingFilters.region}
                      onChange={(selected) => {
                        setPendingFilters(f => ({ ...f, region: selected }));
                        setFilters(f => ({ ...f, region: selected }));
                      }}
                    />
                    <MultiSelect
                      label="Store Name"
                      options={dashboardData.filterOptions.store_name}
                      selected={pendingFilters.store_name}
                      onChange={(selected) => {
                        setPendingFilters(f => ({ ...f, store_name: selected }));
                        setFilters(f => ({ ...f, store_name: selected }));
                      }}
                    />
                    <MultiSelect
                      label="City"
                      options={dashboardData.filterOptions.city}
                      selected={pendingFilters.city}
                      onChange={(selected) => {
                        setPendingFilters(f => ({ ...f, city: selected }));
                        setFilters(f => ({ ...f, city: selected }));
                      }}
                    />
                    <MultiSelect
                      label="Store Format"
                      options={dashboardData.filterOptions.store_format}
                      selected={pendingFilters.store_format}
                      onChange={(selected) => {
                        setPendingFilters(f => ({ ...f, store_format: selected }));
                        setFilters(f => ({ ...f, store_format: selected }));
                      }}
                    />
                    <MultiSelect
                      label="Product Category"
                      options={dashboardData.filterOptions.product_category}
                      selected={pendingFilters.product_category}
                      onChange={(selected) => {
                        setPendingFilters(f => ({ ...f, product_category: selected }));
                        setFilters(f => ({ ...f, product_category: selected }));
                      }}
                    />
                  </div>
                  
                  <div className="flex gap-2 w-full lg:w-auto mt-2 lg:mt-0">
                    <button 
                      onClick={() => { handleApplyFilters(); setIsMobileFiltersOpen(false); }}
                      className="flex-1 lg:flex-none text-sm bg-indigo-500 hover:bg-indigo-600 text-white font-medium px-4 py-2 rounded-lg transition-colors shadow-lg shadow-indigo-500/20"
                    >
                      Apply
                    </button>
                    <button 
                      onClick={handleResetFilters}
                      className="flex-1 lg:flex-none text-sm bg-slate-800 hover:bg-slate-700 text-slate-300 font-medium px-4 py-2 rounded-lg transition-colors border border-slate-700"
                    >
                      Reset
                    </button>
                    <button 
                      onClick={handleExportCSV}
                      className="flex-1 lg:flex-none flex items-center justify-center gap-2 text-sm bg-indigo-500/20 hover:bg-indigo-500/30 text-indigo-300 font-medium px-4 py-2 rounded-lg transition-colors border border-indigo-500/30"
                      title="Export Filtered Data"
                    >
                      <Download className="w-4 h-4" /> <span className="hidden md:inline">Export</span>
                    </button>
                    <button 
                      onClick={handleDownloadInsights}
                      className="flex-1 lg:flex-none flex items-center justify-center gap-2 text-sm bg-white/5 hover:bg-white/10 text-white font-medium px-4 py-2 rounded-lg transition-colors border border-white/10"
                      title="Download Insights"
                    >
                      <FileText className="w-4 h-4" /> <span className="hidden md:inline">Insights</span>
                    </button>
                  </div>
                </div>
              </div>
            </section>

            {/* Charts Section */}
            <div className="flex-1 flex flex-col gap-8 w-full min-w-0">
              
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
                 {/* Row 3: Weekly Net Sales Trend */}
                 <div className="group relative min-h-[400px]">
                    <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
                    <div className="relative h-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col">
                      <h3 className="text-lg font-semibold text-white mb-6">Weekly Net Sales Trend</h3>
                      <div className="flex-1 min-h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={dashboardData.chartDataDate}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" />
                            <XAxis dataKey="date" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} />
                            <YAxis stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} tickFormatter={(val) => `$${val/1000}k`} />
                            <RechartsTooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderColor: '#1e293b', borderRadius: '0.5rem', color: '#f8fafc' }}
                              formatter={(value: number) => formatCurrency(value)}
                            />
                            <Legend wrapperStyle={{ fontSize: '12px', marginTop: '10px' }} />
                            <Line type="monotone" dataKey="revenue" stroke="#818cf8" strokeWidth={3} dot={{ r: 4, fill: '#818cf8' }} activeDot={{ r: 6 }} name="Net Sales" />
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                 </div>

                 {/* Row 3: Sales by Region */}
                 <div className="group relative min-h-[400px]">
                    <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
                    <div className="relative h-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col">
                      <h3 className="text-lg font-semibold text-white mb-6">Sales by Region</h3>
                      <div className="flex-1 min-h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={dashboardData.chartDataRegion}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                            <XAxis dataKey="name" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} angle={-45} textAnchor="end" height={60} />
                            <YAxis stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} tickFormatter={(val) => `$${val/1000}k`} />
                            <RechartsTooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderColor: '#1e293b', borderRadius: '0.5rem', color: '#f8fafc' }}
                              formatter={(value: number) => formatCurrency(value)}
                            />
                            <Bar dataKey="revenue" fill="#38bdf8" radius={[4, 4, 0, 0]} name="Net Sales" />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                 </div>

                 {/* Row 4: Category Performance */}
                 <div className="group relative min-h-[400px]">
                    <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
                    <div className="relative h-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col">
                      <h3 className="text-lg font-semibold text-white mb-6">Product Category Performance</h3>
                      <div className="flex-1 min-h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={dashboardData.chartDataCategory}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                            <XAxis dataKey="name" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} angle={-45} textAnchor="end" height={60} />
                            <YAxis stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} tickFormatter={(val) => `$${val/1000}k`} />
                            <RechartsTooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderColor: '#1e293b', borderRadius: '0.5rem', color: '#f8fafc' }}
                              formatter={(value: number) => formatCurrency(value)}
                            />
                            <Bar dataKey="revenue" fill="#a78bfa" radius={[4, 4, 0, 0]} name="Net Sales" />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                 </div>

                 {/* Row 4: Store Leaderboard */}
                 <div className="group relative min-h-[400px]">
                    <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
                    <div className="relative h-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col">
                      <h3 className="text-lg font-semibold text-white mb-6">Store Leaderboard</h3>
                      <div className="flex-1 min-h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={dashboardData.chartDataStore} layout="vertical" margin={{ left: 10 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" horizontal={true} vertical={false} />
                            <XAxis type="number" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} tickFormatter={(val) => `$${val/1000}k`} />
                            <YAxis dataKey="name" type="category" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} width={120} />
                            <RechartsTooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderColor: '#1e293b', borderRadius: '0.5rem', color: '#f8fafc' }}
                              formatter={(value: number) => formatCurrency(value)}
                            />
                            <Bar dataKey="revenue" fill="#f472b6" radius={[0, 4, 4, 0]} name="Net Sales" />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                 </div>

                 {/* Row 5: Sales Target vs Actual */}
                 <div className="group relative min-h-[400px]">
                    <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
                    <div className="relative h-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col">
                      <h3 className="text-lg font-semibold text-white mb-6">Target vs Actual Sales</h3>
                      <div className="flex-1 min-h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={dashboardData.chartDataDate}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                            <XAxis dataKey="date" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} />
                            <YAxis stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} tickFormatter={(val) => `$${val/1000}k`} />
                            <RechartsTooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderColor: '#1e293b', borderRadius: '0.5rem', color: '#f8fafc' }}
                              formatter={(value: number) => formatCurrency(value)}
                            />
                            <Legend wrapperStyle={{ fontSize: '12px', marginTop: '10px' }} />
                            <Bar dataKey="target" fill="#64748b" radius={[4, 4, 0, 0]} name="Sales Target" />
                            <Bar dataKey="revenue" fill="#38bdf8" radius={[4, 4, 0, 0]} name="Net Sales" />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                 </div>

                 {/* Row 5: Stockout Risk */}
                 <div className="group relative min-h-[400px]">
                    <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
                    <div className="relative h-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col">
                      <h3 className="text-lg font-semibold text-white mb-6">Stockout Risk by Store</h3>
                      <div className="flex-1 min-h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={dashboardData.chartDataStockouts}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                            <XAxis dataKey="name" stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} angle={-45} textAnchor="end" height={60} />
                            <YAxis stroke="#ffffff80" fontSize={11} tickLine={false} axisLine={false} />
                            <RechartsTooltip 
                              contentStyle={{ backgroundColor: '#0f172a', borderColor: '#1e293b', borderRadius: '0.5rem', color: '#f8fafc' }}
                              formatter={(value: number) => formatNumber(value)}
                            />
                            <Bar dataKey="stockouts" fill="#fb7185" radius={[4, 4, 0, 0]} name="Stockouts" />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                 </div>
              </div>
            </div>

            {/* Row 6: Business Insight Summary */}
            <section className="w-full flex-shrink-0 flex flex-col gap-6 mt-2">
              <div className="group relative">
                 <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-purple-500 rounded-2xl opacity-20 group-hover:opacity-30 transition-opacity"></div>
                 <div className="relative bg-white/5 backdrop-blur-xl border border-white/10 p-8 rounded-2xl flex flex-col h-full">
                   <div className="flex items-center gap-3 mb-6">
                     <div className="w-12 h-12 bg-indigo-500/20 rounded-xl flex items-center justify-center">
                       <Sparkles className="w-6 h-6 text-indigo-400" />
                     </div>
                     <h3 className="text-xl font-semibold text-white">Business Insight Summary</h3>
                   </div>
                   
                   <p className="text-base text-indigo-200 mb-6 pb-6 border-b border-white/10">
                     Showing data for <strong className="text-white">{formatNumber(dashboardData.filteredRecordCount)}</strong> records based on the current filter criteria.
                   </p>

                   {dashboardData.insights && dashboardData.insights.length > 0 ? (
                     <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 text-base text-slate-300">
                       {dashboardData.insights.map((insight: string, idx: number) => (
                         <div key={idx} className="flex items-start gap-3 bg-white/5 p-4 rounded-xl border border-white/5">
                           <span className="text-indigo-400 mt-0.5">•</span>
                           <span><strong className="text-indigo-300 font-medium">{insight.split(' ')[0]}</strong> {insight.substring(insight.indexOf(' ') + 1)}</span>
                         </div>
                       ))}
                     </div>
                   ) : (
                     <p className="text-slate-400 italic text-base">No insights available for the current filter criteria.</p>
                   )}
                 </div>
              </div>
            </section>

          </div>
        </main>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 font-sans relative flex flex-col overflow-x-hidden">
      {/* Background Mesh Orbs */}
      <div className="fixed top-[-10%] left-[-10%] w-[500px] h-[500px] bg-blue-600/20 rounded-full blur-[120px] pointer-events-none z-0"></div>
      <div className="fixed bottom-[-10%] right-[-10%] w-[600px] h-[600px] bg-indigo-600/20 rounded-full blur-[150px] pointer-events-none z-0"></div>

      <div className="flex-1 flex flex-col items-center justify-center px-4 py-12 z-10">
        <div className="max-w-4xl w-full">
          {/* Header */}
          <div className="text-center mb-12">
            <h1 className="text-4xl font-extrabold tracking-tight mb-4 text-white">
              Retail Sales Intelligence Dashboard
            </h1>
            <p className="text-slate-400 max-w-xl mx-auto leading-relaxed">
              Upload your operational datasets to synchronize store performance metrics. 
              We use <code className="text-indigo-300 font-mono">store_id</code> to unify your weekly sales with store metadata.
            </p>
          </div>

          {errorMsg && (
            <div className="mb-8 max-w-2xl mx-auto bg-red-500/10 border border-red-500/50 text-red-400 px-4 py-3 rounded-xl flex items-center gap-3">
              <AlertCircle className="w-5 h-5 flex-shrink-0" />
              <p className="font-medium text-sm">{errorMsg}</p>
            </div>
          )}

          {/* Upload Section */}
          <div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full max-w-4xl mx-auto">
            
            {/* Sales Data Upload */}
            <div className="group relative">
              <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
              <label 
                className={`relative bg-white/5 backdrop-blur-xl border rounded-2xl p-8 flex flex-col items-center text-center cursor-pointer transition-all h-full ${
                  salesFile ? 'border-green-500/50 bg-green-500/10' : 'border-white/10 hover:bg-white/10'
                }`}
              >
                {salesFile ? (
                  <div className="flex flex-col items-center justify-center h-full">
                    <CheckCircle2 className="w-12 h-12 text-green-400 mb-4" />
                    <h3 className="text-lg font-semibold text-green-300 mb-1">{salesFile.name}</h3>
                    <p className="text-xs text-green-400/80 mb-6 font-medium">{(salesFile.size / 1024 / 1024).toFixed(2)} MB</p>
                    <div className="w-full py-2.5 bg-white/10 hover:bg-white/20 border border-white/10 rounded-lg text-sm font-medium transition-colors">
                      Change File
                    </div>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full">
                    <div className="w-12 h-12 bg-white/5 rounded-full flex items-center justify-center mb-4">
                      <FileSpreadsheet className="w-6 h-6 text-indigo-400" />
                    </div>
                    <h3 className="text-lg font-semibold text-white mb-1">retail_weekly_sales.xlsx</h3>
                    <p className="text-xs text-slate-500 mb-6">Required: Time-series revenue data</p>
                    <div className="w-full py-2.5 bg-white/10 hover:bg-white/20 border border-white/10 rounded-lg text-sm font-medium transition-colors">
                      Select File
                    </div>
                  </div>
                )}
                <input 
                  type="file" 
                  className="hidden" 
                  accept=".xlsx, .xls"
                  onChange={handleSalesUpload}
                />
              </label>
            </div>

            {/* Store Master Upload */}
            <div className="group relative">
              <div className="absolute -inset-0.5 bg-gradient-to-br from-indigo-500 to-blue-500 rounded-2xl opacity-20 group-hover:opacity-40 transition-opacity"></div>
              <label 
                className={`relative bg-white/5 backdrop-blur-xl border rounded-2xl p-8 flex flex-col items-center text-center cursor-pointer transition-all h-full ${
                  storeFile ? 'border-green-500/50 bg-green-500/10' : 'border-white/10 hover:bg-white/10'
                }`}
              >
                {storeFile ? (
                  <div className="flex flex-col items-center justify-center h-full">
                    <CheckCircle2 className="w-12 h-12 text-green-400 mb-4" />
                    <h3 className="text-lg font-semibold text-green-300 mb-1">{storeFile.name}</h3>
                    <p className="text-xs text-green-400/80 mb-6 font-medium">{(storeFile.size / 1024 / 1024).toFixed(2)} MB</p>
                    <div className="w-full py-2.5 bg-white/10 hover:bg-white/20 border border-white/10 rounded-lg text-sm font-medium transition-colors">
                      Change File
                    </div>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center h-full">
                    <div className="w-12 h-12 bg-white/5 rounded-full flex items-center justify-center mb-4">
                      <Store className="w-6 h-6 text-indigo-400" />
                    </div>
                    <h3 className="text-lg font-semibold text-white mb-1">store_master.xlsx</h3>
                    <p className="text-xs text-slate-500 mb-6">Required: Location & Hierarchy metadata</p>
                    <div className="w-full py-2.5 bg-white/10 hover:bg-white/20 border border-white/10 rounded-lg text-sm font-medium transition-colors">
                      Select File
                    </div>
                  </div>
                )}
                <input 
                  type="file" 
                  className="hidden" 
                  accept=".xlsx, .xls"
                  onChange={handleStoreUpload}
                />
              </label>
            </div>
          </div>

          <div className="mt-12 max-w-md mx-auto">
            <button
              onClick={handleLoadDashboard}
              disabled={!salesFile || !storeFile || isLoading}
              className={`w-full py-4 rounded-xl font-bold shadow-lg flex items-center justify-center gap-2 transition-all ${
                salesFile && storeFile && !isLoading
                  ? 'bg-indigo-600 hover:bg-indigo-500 text-white shadow-indigo-900/40' 
                  : 'bg-white/5 text-slate-500 cursor-not-allowed border border-white/10'
              }`}
            >
              {isLoading ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  <span>Processing Data...</span>
                </>
              ) : (
                <>
                  <span>Generate Dashboard</span>
                  <ArrowRight className={`w-5 h-5 ${salesFile && storeFile ? 'opacity-100' : 'opacity-50'}`} />
                </>
              )}
            </button>
            <p className="text-center text-[10px] text-slate-500 mt-4 tracking-widest uppercase">
              No data is stored externally. Processing occurs in-memory.
            </p>
          </div>
        </div>
        </div>
      </div>
    </div>
  );
}


