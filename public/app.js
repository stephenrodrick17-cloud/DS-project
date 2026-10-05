/**
 * NexusStore — Live Excel-Driven E-Commerce Application
 * Single Source of Truth: products_database.xlsx (Sheet: "Products")
 * Features:
 * - Real-time 5-second polling of Excel database
 * - Dynamic image catalog with curated high-res product photography
 * - Hash router: #catalog, #product/:id, #cart, #checkout, #order-success, #sync-status
 * - Change detection with highlight animation for modified Excel values
 * - Persistent cart with promo code support (EXCEL10)
 * - Working checkout and order placement via /api/orders
 * - Slide-out drawer and full page views
 */

// ============================================================================
// Curated Product Images & Fallbacks
// ============================================================================
const PRODUCT_IMAGES = {
  1: 'https://images.unsplash.com/photo-1527864550417-7fd91fc51a46?w=700&auto=format&fit=crop&q=80', // Wireless Mouse
  2: 'https://images.unsplash.com/photo-1587829741301-dc798b83add3?w=700&auto=format&fit=crop&q=80', // Mechanical Keyboard
  3: 'https://images.unsplash.com/photo-1583863788434-e58a36330cf0?w=700&auto=format&fit=crop&q=80', // USB-C Charger 65W
  4: 'https://images.unsplash.com/photo-1608043152269-423dbba4e7e1?w=700&auto=format&fit=crop&q=80', // Bluetooth Speaker
  5: 'https://images.unsplash.com/photo-1521572267360-ee0c2909d518?w=700&auto=format&fit=crop&q=80', // Cotton T-Shirt
  6: 'https://images.unsplash.com/photo-1576995853123-5a10305d93c0?w=700&auto=format&fit=crop&q=80', // Denim Jacket
  7: 'https://images.unsplash.com/photo-1542291026-7eec264c27ff?w=700&auto=format&fit=crop&q=80', // Running Shoes
  8: 'https://images.unsplash.com/photo-1627123424574-724758594e93?w=700&auto=format&fit=crop&q=80', // Leather Wallet
  9: 'https://images.unsplash.com/photo-1602143407151-7111542de6e8?w=700&auto=format&fit=crop&q=80', // Steel Water Bottle
  10: 'https://images.unsplash.com/photo-1534349762230-e0cadf78f5da?w=700&auto=format&fit=crop&q=80', // Desk Lamp
  11: 'https://images.unsplash.com/photo-1544716278-ca5e3f4abd8c?w=700&auto=format&fit=crop&q=80', // Notebook (Pack of 5)
  12: 'https://images.unsplash.com/photo-1585336261026-70e1c668631b?w=700&auto=format&fit=crop&q=80', // Gel Pen Set
  13: 'https://images.unsplash.com/photo-1588702547919-26089e690ecc?w=700&auto=format&fit=crop&q=80'  // 4K Pro Webcam
};

const CATEGORY_FALLBACK_IMAGES = {
  electronics: 'https://images.unsplash.com/photo-1505740420928-5e560c06d30e?w=700&auto=format&fit=crop&q=80',
  clothing: 'https://images.unsplash.com/photo-1489987707025-afc232f7ea0f?w=700&auto=format&fit=crop&q=80',
  footwear: 'https://images.unsplash.com/photo-1560769629-975ec94e6a86?w=700&auto=format&fit=crop&q=80',
  accessories: 'https://images.unsplash.com/photo-1523275335684-37898b6baf30?w=700&auto=format&fit=crop&q=80',
  home: 'https://images.unsplash.com/photo-1513694203232-719a280e022f?w=700&auto=format&fit=crop&q=80',
  stationery: 'https://images.unsplash.com/photo-1456735190827-d1262f71b8a3?w=700&auto=format&fit=crop&q=80',
  default: 'https://images.unsplash.com/photo-1526170375885-4d8ecf77b99f?w=700&auto=format&fit=crop&q=80'
};

function getProductImage(product) {
  if (product.Image && typeof product.Image === 'string' && product.Image.startsWith('http')) {
    return product.Image;
  }
  const idNum = Number(product.ID);
  if (PRODUCT_IMAGES[idNum]) {
    return PRODUCT_IMAGES[idNum];
  }
  const cat = String(product.Category || '').toLowerCase();
  return CATEGORY_FALLBACK_IMAGES[cat] || CATEGORY_FALLBACK_IMAGES.default;
}

// ============================================================================
// Application State
// ============================================================================
const state = {
  products: [],
  previousProductsMap: new Map(),
  cart: loadCartFromStorage(),
  discountPercent: 0,
  activePromoCode: '',
  selectedCategory: 'all',
  selectedStatus: 'all',
  selectedSort: 'id-asc',
  searchQuery: '',
  activeView: 'grid', // 'grid' | 'table'
  currentPage: 'catalog',
  currentProductId: null,
  lastOrder: null,
  pollIntervalSec: 5,
  countdown: 5,
  isInitialLoad: true,
  timerId: null,
  countdownTimerId: null
};

// ============================================================================
// DOM Elements Cache
// ============================================================================
const dom = {
  pages: {
    catalog: document.getElementById('page-catalog'),
    product: document.getElementById('page-product'),
    cart: document.getElementById('page-cart'),
    checkout: document.getElementById('page-checkout'),
    orderSuccess: document.getElementById('page-order-success'),
    syncStatus: document.getElementById('page-sync-status')
  },
  navLinks: document.querySelectorAll('.nav-link'),
  navCartBadge: document.getElementById('nav-cart-badge'),
  cartToggleBtn: document.getElementById('cart-toggle-btn'),
  refreshBtn: document.getElementById('refresh-btn'),
  refreshIcon: document.getElementById('refresh-icon'),
  connectionStatus: document.getElementById('connection-status'),
  countdownBadge: document.getElementById('countdown-badge'),
  liveIndicator: document.getElementById('live-indicator'),
  errorBanner: document.getElementById('error-banner'),
  errorMessage: document.getElementById('error-message'),
  dismissError: document.getElementById('dismiss-error'),
  toastContainer: document.getElementById('toast-container'),
  // Catalog page elements
  statTotal: document.getElementById('stat-total'),
  statInStock: document.getElementById('stat-in-stock'),
  statLowStock: document.getElementById('stat-low-stock'),
  statOutStock: document.getElementById('stat-out-stock'),
  categoryChips: document.getElementById('category-chips'),
  searchInput: document.getElementById('search-input'),
  clearSearchBtn: document.getElementById('clear-search'),
  statusFilter: document.getElementById('status-filter'),
  sortSelect: document.getElementById('sort-select'),
  viewGridBtn: document.getElementById('view-grid-btn'),
  viewTableBtn: document.getElementById('view-table-btn'),
  resultsCount: document.getElementById('results-count'),
  resetFiltersBtn: document.getElementById('reset-filters-btn'),
  emptyResetBtn: document.getElementById('empty-reset-btn'),
  gridView: document.getElementById('grid-view'),
  tableView: document.getElementById('table-view'),
  tableBody: document.getElementById('table-body'),
  emptyState: document.getElementById('empty-state'),
  // Product Detail Page (PDP)
  pdpImage: document.getElementById('pdp-image'),
  pdpStatusBadge: document.getElementById('pdp-status-badge'),
  pdpBreadcrumbCat: document.getElementById('pdp-breadcrumb-cat'),
  pdpBreadcrumbName: document.getElementById('pdp-breadcrumb-name'),
  pdpCategory: document.getElementById('pdp-category'),
  pdpId: document.getElementById('pdp-id'),
  pdpTitle: document.getElementById('pdp-title'),
  pdpStars: document.getElementById('pdp-stars'),
  pdpRatingNum: document.getElementById('pdp-rating-num'),
  pdpPrice: document.getElementById('pdp-price'),
  pdpMrp: document.getElementById('pdp-mrp'),
  pdpDiscount: document.getElementById('pdp-discount'),
  pdpStockCount: document.getElementById('pdp-stock-count'),
  pdpStatusText: document.getElementById('pdp-status-text'),
  pdpStockFill: document.getElementById('pdp-stock-fill'),
  pdpQtyInput: document.getElementById('pdp-qty-input'),
  qtyMinus: document.getElementById('qty-minus'),
  qtyPlus: document.getElementById('qty-plus'),
  qtyMaxNote: document.getElementById('qty-max-note'),
  pdpAddCartBtn: document.getElementById('pdp-add-cart-btn'),
  pdpBuyNowBtn: document.getElementById('pdp-buy-now-btn'),
  // Cart Page
  cartContentWrapper: document.getElementById('cart-content-wrapper'),
  cartItemsList: document.getElementById('cart-items-list'),
  cartEmptyView: document.getElementById('cart-empty-view'),
  cartSubtotal: document.getElementById('cart-subtotal'),
  cartShipping: document.getElementById('cart-shipping'),
  cartTax: document.getElementById('cart-tax'),
  discountRow: document.getElementById('discount-row'),
  cartDiscount: document.getElementById('cart-discount'),
  couponInput: document.getElementById('coupon-input'),
  applyCouponBtn: document.getElementById('apply-coupon-btn'),
  couponMessage: document.getElementById('coupon-message'),
  cartGrandTotal: document.getElementById('cart-grand-total'),
  cartProceedCheckoutBtn: document.getElementById('cart-proceed-checkout-btn'),
  clearCartBtn: document.getElementById('clear-cart-btn'),
  // Checkout Page
  checkoutForm: document.getElementById('checkout-form'),
  placeOrderBtn: document.getElementById('place-order-btn'),
  checkoutItemsPreview: document.getElementById('checkout-items-preview'),
  checkoutSubtotal: document.getElementById('checkout-subtotal'),
  checkoutTax: document.getElementById('checkout-tax'),
  checkoutDiscountRow: document.getElementById('checkout-discount-row'),
  checkoutDiscount: document.getElementById('checkout-discount'),
  checkoutGrandTotal: document.getElementById('checkout-grand-total'),
  // Order Success Page
  successOrderId: document.getElementById('success-order-id'),
  successDeliveryDate: document.getElementById('success-delivery-date'),
  successItemsList: document.getElementById('success-items-list'),
  successPaymentMethod: document.getElementById('success-payment-method'),
  successCustomerName: document.getElementById('success-customer-name'),
  successTotalPaid: document.getElementById('success-total-paid'),
  successTimestamp: document.getElementById('success-timestamp'),
  // Drawer
  drawerOverlay: document.getElementById('cart-drawer-overlay'),
  closeDrawerBtn: document.getElementById('close-drawer-btn'),
  drawerItemsCount: document.getElementById('drawer-items-count'),
  drawerItemsList: document.getElementById('drawer-items-list'),
  drawerSubtotal: document.getElementById('drawer-subtotal'),
  drawerCheckoutBtn: document.getElementById('drawer-checkout-btn'),
  // Admin Live Monitor
  adminRowCount: document.getElementById('admin-row-count'),
  adminLastModified: document.getElementById('admin-last-modified'),
  adminResetDbBtn: document.getElementById('admin-reset-db-btn'),
  // Guide Modal
  guideModal: document.getElementById('guide-modal'),
  topGuideBtn: document.getElementById('top-guide-btn'),
  closeGuideBtn: document.getElementById('close-guide-btn'),
  gotItBtn: document.getElementById('got-it-btn'),
  // Razorpay Checkout Modal
  rzpModal: document.getElementById('razorpay-modal'),
  rzpAmountDisplay: document.getElementById('rzp-amount-display'),
  rzpOrderPreviewId: document.getElementById('rzp-order-preview-id'),
  rzpCloseBtn: document.getElementById('rzp-close-btn'),
  rzpProcessingOverlay: document.getElementById('rzp-processing-overlay'),
  rzpSimulateSuccessBtn: document.getElementById('rzp-simulate-success-btn'),
  rzpSimulateFailureBtn: document.getElementById('rzp-simulate-failure-btn'),
  rzpUpiVerifyPayBtn: document.getElementById('rzp-upi-verify-pay-btn'),
  rzpCardPayBtn: document.getElementById('rzp-card-pay-btn'),
  rzpNetbankingPayBtn: document.getElementById('rzp-netbanking-pay-btn'),
  rzpWalletPayBtn: document.getElementById('rzp-wallet-pay-btn'),
  rzpQrTimerText: document.getElementById('rzp-qr-timer-text'),
  successRazorpayIdRow: document.getElementById('success-razorpay-id-row'),
  successRazorpayId: document.getElementById('success-razorpay-id'),
  successPaymentStatus: document.getElementById('success-payment-status')
};

// ============================================================================
// Initialization & Router Setup
// ============================================================================
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  updateCartBadge();
  window.addEventListener('hashchange', handleRouting);
  handleRouting(); // Initial route
  fetchProducts(); // Initial fetch
  startAutoRefresh();
});

/**
 * Hash-based Client Router
 * Supported paths:
 * - #catalog (or empty) -> Shop Catalog
 * - #product/:id -> Product Details Page
 * - #cart -> Shopping Cart
 * - #checkout -> Checkout
 * - #order-success -> Order Confirmation
 * - #sync-status -> Excel Live Monitor
 */
function handleRouting() {
  const hash = window.location.hash || '#catalog';
  closeCartDrawer();

  if (hash.startsWith('#product/')) {
    const id = hash.replace('#product/', '');
    state.currentProductId = id;
    showPage('product');
    renderProductDetailPage(id);
  } else if (hash === '#cart') {
    showPage('cart');
    renderCartPage();
  } else if (hash === '#checkout') {
    showPage('checkout');
    renderCheckoutPage();
  } else if (hash.startsWith('#order-success')) {
    showPage('orderSuccess');
    renderOrderSuccessPage();
  } else if (hash === '#sync-status') {
    showPage('syncStatus');
    renderAdminPage();
  } else {
    showPage('catalog');
  }

  // Update active nav links
  dom.navLinks.forEach(link => {
    const pageTarget = link.getAttribute('href');
    link.classList.toggle('active', pageTarget === hash || (hash === '#catalog' && pageTarget === '#catalog'));
  });

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function showPage(pageKey) {
  state.currentPage = pageKey;
  Object.keys(dom.pages).forEach(key => {
    if (dom.pages[key]) {
      dom.pages[key].classList.toggle('hidden', key !== pageKey);
      dom.pages[key].classList.toggle('active', key === pageKey);
    }
  });
}

function navigateTo(hash) {
  window.location.hash = hash;
}

// ============================================================================
// Event Listeners Setup
// ============================================================================
function setupEventListeners() {
  // Search input
  let searchTimeout;
  dom.searchInput.addEventListener('input', (e) => {
    clearTimeout(searchTimeout);
    state.searchQuery = e.target.value.trim().toLowerCase();
    dom.clearSearchBtn.classList.toggle('hidden', state.searchQuery === '');
    searchTimeout = setTimeout(() => renderCatalog(), 150);
  });

  dom.clearSearchBtn.addEventListener('click', () => {
    dom.searchInput.value = '';
    state.searchQuery = '';
    dom.clearSearchBtn.classList.add('hidden');
    renderCatalog();
  });

  // Filters & Sorting
  dom.statusFilter.addEventListener('change', (e) => {
    state.selectedStatus = e.target.value;
    renderCatalog();
  });

  dom.sortSelect.addEventListener('change', (e) => {
    state.selectedSort = e.target.value;
    renderCatalog();
  });

  // View switchers
  dom.viewGridBtn.addEventListener('click', () => setView('grid'));
  dom.viewTableBtn.addEventListener('click', () => setView('table'));

  // Reset filters
  dom.resetFiltersBtn.addEventListener('click', resetFilters);
  dom.emptyResetBtn.addEventListener('click', resetFilters);

  // Manual Refresh
  dom.refreshBtn.addEventListener('click', () => {
    dom.refreshIcon.classList.add('spinning');
    fetchProducts(true).finally(() => {
      setTimeout(() => dom.refreshIcon.classList.remove('spinning'), 500);
    });
    resetCountdown();
  });

  // Cart Drawer open/close
  dom.cartToggleBtn.addEventListener('click', openCartDrawer);
  dom.closeDrawerBtn.addEventListener('click', closeCartDrawer);
  dom.drawerOverlay.addEventListener('click', (e) => {
    if (e.target === dom.drawerOverlay) closeCartDrawer();
  });

  // Guide Modal
  dom.topGuideBtn.addEventListener('click', () => dom.guideModal.classList.remove('hidden'));
  dom.closeGuideBtn.addEventListener('click', () => dom.guideModal.classList.add('hidden'));
  dom.gotItBtn.addEventListener('click', () => dom.guideModal.classList.add('hidden'));
  dom.guideModal.addEventListener('click', (e) => {
    if (e.target === dom.guideModal) dom.guideModal.classList.add('hidden');
  });

  // Error Banner dismiss
  dom.dismissError.addEventListener('click', () => dom.errorBanner.classList.add('hidden'));

  // Coupon application
  dom.applyCouponBtn.addEventListener('click', applyPromoCode);

  // Clear Cart
  dom.clearCartBtn.addEventListener('click', () => {
    if (confirm('Are you sure you want to empty your cart?')) {
      clearCart();
      renderCartPage();
    }
  });

  // PDP Quantity buttons
  dom.qtyMinus.addEventListener('click', () => {
    let cur = parseInt(dom.pdpQtyInput.value, 10) || 1;
    if (cur > 1) dom.pdpQtyInput.value = cur - 1;
  });

  dom.qtyPlus.addEventListener('click', () => {
    const product = state.products.find(p => String(p.ID) === String(state.currentProductId));
    const maxStock = product ? (Number(product.Stock) || 99) : 99;
    let cur = parseInt(dom.pdpQtyInput.value, 10) || 1;
    if (cur < maxStock) dom.pdpQtyInput.value = cur + 1;
  });

  // PDP Add to Cart / Buy Now
  dom.pdpAddCartBtn.addEventListener('click', () => {
    if (!state.currentProductId) return;
    const qty = parseInt(dom.pdpQtyInput.value, 10) || 1;
    addToCart(state.currentProductId, qty);
    openCartDrawer();
  });

  dom.pdpBuyNowBtn.addEventListener('click', () => {
    if (!state.currentProductId) return;
    const qty = parseInt(dom.pdpQtyInput.value, 10) || 1;
    addToCart(state.currentProductId, qty);
    navigateTo('#checkout');
  });

  // Proceed to Checkout button on Cart Page
  dom.cartProceedCheckoutBtn.addEventListener('click', () => {
    if (state.cart.length === 0) {
      alert('Your cart is empty.');
      return;
    }
    navigateTo('#checkout');
  });

  // Checkout Payment method selector
  document.querySelectorAll('input[name="payment-method"]').forEach(radio => {
    radio.addEventListener('change', () => {
      document.querySelectorAll('.payment-method-card').forEach(c => c.classList.remove('active'));
      radio.closest('.payment-method-card').classList.add('active');
    });
  });

  // Place Order form submit
  dom.checkoutForm.addEventListener('submit', handlePlaceOrder);

  // Admin Reset Database
  dom.adminResetDbBtn.addEventListener('click', async () => {
    if (!confirm('Reset products_database.xlsx back to the default 12 products?')) return;
    try {
      const res = await fetch('/api/reset-database', { method: 'POST' });
      const data = await res.json();
      if (data.success) {
        showToast('Database Reset', 'products_database.xlsx has been reset to default.');
        await fetchProducts(true);
      }
    } catch (e) {
      alert('Failed to reset: ' + e.message);
    }
  });

  // Setup Razorpay Modal Listeners
  setupRazorpayModalListeners();
}

function setView(view) {
  state.activeView = view;
  dom.viewGridBtn.classList.toggle('active', view === 'grid');
  dom.viewTableBtn.classList.toggle('active', view === 'table');
  dom.gridView.classList.toggle('hidden', view !== 'grid');
  dom.tableView.classList.toggle('hidden', view !== 'table');
}

// ============================================================================
// Auto-Refresh & Background Polling (5 Seconds)
// ============================================================================
function startAutoRefresh() {
  if (state.timerId) clearInterval(state.timerId);
  if (state.countdownTimerId) clearInterval(state.countdownTimerId);

  state.countdown = state.pollIntervalSec;
  updateCountdownBadge();

  state.countdownTimerId = setInterval(() => {
    state.countdown -= 1;
    if (state.countdown <= 0) state.countdown = state.pollIntervalSec;
    updateCountdownBadge();
  }, 1000);

  state.timerId = setInterval(() => {
    fetchProducts();
  }, state.pollIntervalSec * 1000);
}

function resetCountdown() {
  state.countdown = state.pollIntervalSec;
  updateCountdownBadge();
}

function updateCountdownBadge() {
  dom.countdownBadge.textContent = `${state.countdown}s`;
}

// ============================================================================
// Fetch Data from Excel API & Change Detection
// ============================================================================
async function fetchProducts(isManual = false) {
  try {
    const res = await fetch('/api/products', {
      headers: { 'Accept': 'application/json', 'Cache-Control': 'no-cache' }
    });

    if (!res.ok) {
      const errData = await res.json().catch(() => ({}));
      throw new Error(errData.error || `Server error: ${res.status}`);
    }

    const data = await res.json();
    if (!Array.isArray(data)) throw new Error('Expected array of products');

    dom.errorBanner.classList.add('hidden');
    dom.liveIndicator.classList.remove('error');
    dom.connectionStatus.textContent = 'Live';

    // Detect what changed in Excel since last fetch
    const changesMap = detectExcelChanges(data);

    state.products = data;
    updateHeroStats(data);
    renderCategoryChips(data);

    // Re-render whichever page is active
    if (state.currentPage === 'catalog') {
      renderCatalog(changesMap);
    } else if (state.currentPage === 'product' && state.currentProductId) {
      renderProductDetailPage(state.currentProductId, changesMap);
    } else if (state.currentPage === 'syncStatus') {
      renderAdminPage();
    }

    // If changes occurred, notify the user with a sleek toast
    if (changesMap.size > 0 && !state.isInitialLoad) {
      notifyExcelChanges(changesMap);
    }

    // Save snapshot
    const map = new Map();
    data.forEach(p => map.set(String(p.ID), { ...p }));
    state.previousProductsMap = map;
    state.isInitialLoad = false;

  } catch (err) {
    console.warn('[Sync Error]:', err.message);
    dom.liveIndicator.classList.add('error');
    dom.connectionStatus.textContent = 'Sync Issue';
    dom.errorBanner.classList.remove('hidden');
    dom.errorMessage.textContent = `${err.message} (Auto-retrying in 5s...)`;
  }
}

function detectExcelChanges(freshProducts) {
  const changes = new Map();
  if (state.isInitialLoad || state.previousProductsMap.size === 0) return changes;

  for (const fresh of freshProducts) {
    const id = String(fresh.ID);
    const old = state.previousProductsMap.get(id);

    if (!old) {
      changes.set(id, {
        isNew: true,
        product: fresh,
        fields: new Set(['Name', 'Category', 'Price (INR)', 'Stock', 'Rating', 'Status'])
      });
      continue;
    }

    const diffFields = new Set();
    const details = [];
    const fieldsToInspect = ['Name', 'Category', 'Price (INR)', 'Stock', 'Rating', 'Status'];

    for (const f of fieldsToInspect) {
      if (String(old[f]) !== String(fresh[f])) {
        diffFields.add(f);
        details.push({ field: f, oldVal: old[f], newVal: fresh[f] });
      }
    }

    if (diffFields.size > 0) {
      changes.set(id, {
        isNew: false,
        product: fresh,
        fields: diffFields,
        details
      });
    }
  }

  return changes;
}

function notifyExcelChanges(changesMap) {
  for (const [id, info] of changesMap.entries()) {
    const p = info.product;
    if (info.isNew) {
      showToast('New Product Added in Excel!', `<strong>${escapeHtml(p.Name)}</strong> (ID #${p.ID}) is now live.`);
    } else {
      const summary = info.details
        .map(d => `<div>&bull; ${d.field}: <s>${d.oldVal}</s> ➔ <span style="color:#fbbf24">${d.newVal}</span></div>`)
        .join('');
      showToast('Excel Value Updated!', `<strong>${escapeHtml(p.Name)}</strong><br>${summary}`);
    }
  }
}

function showToast(title, htmlBody) {
  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerHTML = `
    <div class="toast-icon">⚡</div>
    <div>
      <div class="toast-title">${title}</div>
      <div class="toast-body">${htmlBody}</div>
    </div>
  `;
  dom.toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(-10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 4500);
}

// ============================================================================
// PAGE 1: Catalog Rendering & Controls
// ============================================================================
function updateHeroStats(products) {
  const total = products.length;
  let inStock = 0, lowStock = 0, outStock = 0;

  products.forEach(p => {
    const s = String(p.Status || '').toLowerCase();
    if (s === 'in stock') inStock++;
    else if (s === 'low stock') lowStock++;
    else if (s === 'out of stock') outStock++;
  });

  dom.statTotal.textContent = total;
  dom.statInStock.textContent = inStock;
  dom.statLowStock.textContent = lowStock;
  dom.statOutStock.textContent = outStock;
}

function renderCategoryChips(products) {
  const categories = [...new Set(products.map(p => p.Category).filter(Boolean))].sort();
  const current = state.selectedCategory;

  let html = `<button class="chip ${current === 'all' ? 'active' : ''}" data-cat="all">All Products (${products.length})</button>`;
  categories.forEach(cat => {
    const count = products.filter(p => p.Category === cat).length;
    html += `<button class="chip ${current === cat ? 'active' : ''}" data-cat="${escapeHtml(cat)}">${escapeHtml(cat)} (${count})</button>`;
  });

  dom.categoryChips.innerHTML = html;

  dom.categoryChips.querySelectorAll('.chip').forEach(chip => {
    chip.addEventListener('click', () => {
      dom.categoryChips.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      state.selectedCategory = chip.dataset.cat;
      renderCatalog();
    });
  });
}

function getFilteredAndSortedProducts() {
  let list = [...state.products];

  // Search
  if (state.searchQuery) {
    const q = state.searchQuery;
    list = list.filter(p => {
      return (p.Name || '').toLowerCase().includes(q) ||
             (p.Category || '').toLowerCase().includes(q) ||
             String(p.ID || '').toLowerCase().includes(q);
    });
  }

  // Category
  if (state.selectedCategory !== 'all') {
    list = list.filter(p => p.Category === state.selectedCategory);
  }

  // Status
  if (state.selectedStatus !== 'all') {
    list = list.filter(p => (p.Status || '').toLowerCase() === state.selectedStatus.toLowerCase());
  }

  // Sorting
  list.sort((a, b) => {
    switch (state.selectedSort) {
      case 'price-asc': return (Number(a['Price (INR)']) || 0) - (Number(b['Price (INR)']) || 0);
      case 'price-desc': return (Number(b['Price (INR)']) || 0) - (Number(a['Price (INR)']) || 0);
      case 'rating-desc': return (Number(b.Rating) || 0) - (Number(a.Rating) || 0);
      case 'stock-asc': return (Number(a.Stock) || 0) - (Number(b.Stock) || 0);
      case 'stock-desc': return (Number(b.Stock) || 0) - (Number(a.Stock) || 0);
      case 'name-asc': return String(a.Name).localeCompare(String(b.Name));
      case 'id-asc':
      default: return (Number(a.ID) || 0) - (Number(b.ID) || 0);
    }
  });

  return list;
}

function renderCatalog(changesMap = new Map()) {
  const filtered = getFilteredAndSortedProducts();
  const hasFilters = state.searchQuery !== '' || state.selectedCategory !== 'all' || state.selectedStatus !== 'all';

  dom.resultsCount.textContent = hasFilters
    ? `Showing ${filtered.length} of ${state.products.length} products`
    : `Showing all ${state.products.length} products`;
  dom.resetFiltersBtn.classList.toggle('hidden', !hasFilters);

  if (filtered.length === 0) {
    dom.gridView.innerHTML = '';
    dom.tableBody.innerHTML = '';
    dom.emptyState.classList.remove('hidden');
    return;
  }

  dom.emptyState.classList.add('hidden');
  renderGridView(filtered, changesMap);
  renderTableView(filtered, changesMap);
}

function renderGridView(products, changesMap) {
  dom.gridView.innerHTML = products.map(p => {
    const id = String(p.ID);
    const changeInfo = changesMap.get(id);
    const changedFields = changeInfo ? changeInfo.fields : new Set();
    const isCardChanged = changedFields.size > 0;

    const imgUrl = getProductImage(p);
    const statusClass = getStatusClass(p.Status);
    const isOutOfStock = String(p.Status || '').toLowerCase() === 'out of stock' || Number(p.Stock) <= 0;
    const priceFormatted = Number(p['Price (INR)']).toLocaleString('en-IN');
    const originalPrice = Math.round(Number(p['Price (INR)']) * 1.3).toLocaleString('en-IN');

    const nameClass = changedFields.has('Name') ? 'value-changed' : '';
    const priceClass = changedFields.has('Price (INR)') ? 'value-changed' : '';
    const stockClass = changedFields.has('Stock') ? 'value-changed' : '';
    const statusClassFlash = changedFields.has('Status') ? 'value-changed' : '';

    return `
      <article class="product-card ${isCardChanged ? 'card-changed' : ''}" data-id="${escapeHtml(id)}">
        <div class="card-image-box" onclick="navigateTo('#product/${escapeHtml(id)}')">
          <img src="${imgUrl}" alt="${escapeHtml(p.Name)}" class="card-img" loading="lazy" onerror="this.src='${CATEGORY_FALLBACK_IMAGES.default}'">
          <span class="card-badge-status status-badge ${statusClass} ${statusClassFlash}">${escapeHtml(p.Status || 'In Stock')}</span>
          <span class="card-badge-id">#${escapeHtml(id)}</span>
        </div>

        <div class="card-body">
          <div class="card-category">${escapeHtml(p.Category || 'General')}</div>
          <h2 class="card-title ${nameClass}" onclick="navigateTo('#product/${escapeHtml(id)}')">${escapeHtml(p.Name)}</h2>

          <div class="card-rating-row">
            <span class="rating-stars">${renderStars(p.Rating)}</span>
            <span class="rating-score">${Number(p.Rating).toFixed(1)}</span>
            <span class="stock-tag-text ${stockClass}">&bull; Stock: ${Number(p.Stock) || 0}</span>
          </div>

          <div class="card-price-row">
            <div class="price-main">
              <span class="price-inr ${priceClass}">₹${priceFormatted}</span>
              <span class="price-strikethrough">₹${originalPrice}</span>
            </div>
          </div>

          <div class="card-actions">
            <button class="btn-card-details" onclick="navigateTo('#product/${escapeHtml(id)}')">View Details</button>
            <button class="btn-card-cart" onclick="addToCart('${escapeHtml(id)}', 1)" ${isOutOfStock ? 'disabled title="Currently Out of Stock"' : ''}>
              ${isOutOfStock ? 'Out of Stock' : 'Add to Cart'}
            </button>
          </div>
        </div>
      </article>
    `;
  }).join('');
}

function renderTableView(products, changesMap) {
  dom.tableBody.innerHTML = products.map(p => {
    const id = String(p.ID);
    const changeInfo = changesMap.get(id);
    const changedFields = changeInfo ? changeInfo.fields : new Set();
    const isRowChanged = changedFields.size > 0;

    const imgUrl = getProductImage(p);
    const statusClass = getStatusClass(p.Status);
    const isOutOfStock = String(p.Status || '').toLowerCase() === 'out of stock' || Number(p.Stock) <= 0;
    const priceFormatted = Number(p['Price (INR)']).toLocaleString('en-IN');

    const nameClass = changedFields.has('Name') ? 'value-changed' : '';
    const priceClass = changedFields.has('Price (INR)') ? 'value-changed' : '';
    const stockClass = changedFields.has('Stock') ? 'value-changed' : '';
    const statusClassFlash = changedFields.has('Status') ? 'value-changed' : '';

    return `
      <tr class="${isRowChanged ? 'row-changed' : ''}">
        <td><img src="${imgUrl}" alt="${escapeHtml(p.Name)}" class="table-thumb" onerror="this.src='${CATEGORY_FALLBACK_IMAGES.default}'"></td>
        <td><code>#${escapeHtml(id)}</code></td>
        <td class="${nameClass}"><strong>${escapeHtml(p.Name)}</strong></td>
        <td><span class="chip">${escapeHtml(p.Category || 'General')}</span></td>
        <td class="${priceClass}"><strong>₹${priceFormatted}</strong></td>
        <td class="${stockClass}">${Number(p.Stock) || 0}</td>
        <td>★ ${Number(p.Rating).toFixed(1)}</td>
        <td><span class="status-badge ${statusClass} ${statusClassFlash}">${escapeHtml(p.Status || 'In Stock')}</span></td>
        <td>
          <button class="btn-primary" style="padding: 0.35rem 0.75rem; font-size: 0.75rem;" onclick="navigateTo('#product/${escapeHtml(id)}')">
            Buy
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function resetFilters() {
  state.searchQuery = '';
  state.selectedCategory = 'all';
  state.selectedStatus = 'all';
  state.selectedSort = 'id-asc';

  dom.searchInput.value = '';
  dom.clearSearchBtn.classList.add('hidden');
  dom.statusFilter.value = 'all';
  dom.sortSelect.value = 'id-asc';

  renderCategoryChips(state.products);
  renderCatalog();
}

// ============================================================================
// PAGE 2: Product Detail Page (PDP)
// ============================================================================
function renderProductDetailPage(id, changesMap = new Map()) {
  const product = state.products.find(p => String(p.ID) === String(id));

  if (!product) {
    // If not found yet, show skeleton or redirect
    return;
  }

  const changeInfo = changesMap.get(String(id));
  const changedFields = changeInfo ? changeInfo.fields : new Set();

  const imgUrl = getProductImage(product);
  const statusClass = getStatusClass(product.Status);
  const isOutOfStock = String(product.Status || '').toLowerCase() === 'out of stock' || Number(product.Stock) <= 0;
  const price = Number(product['Price (INR)']) || 0;
  const originalPrice = Math.round(price * 1.3);
  const discountPercent = Math.round(((originalPrice - price) / originalPrice) * 100);
  const stock = Number(product.Stock) || 0;
  const stockPercent = Math.min(100, Math.max(0, (stock / 100) * 100));

  // Populate elements
  dom.pdpImage.src = imgUrl;
  dom.pdpImage.onerror = () => { dom.pdpImage.src = CATEGORY_FALLBACK_IMAGES.default; };

  dom.pdpBreadcrumbCat.textContent = product.Category || 'General';
  dom.pdpBreadcrumbName.textContent = product.Name;

  dom.pdpCategory.textContent = product.Category || 'General';
  dom.pdpId.textContent = `ID #${product.ID}`;

  dom.pdpTitle.textContent = product.Name;
  if (changedFields.has('Name')) dom.pdpTitle.classList.add('value-changed');

  dom.pdpStars.innerHTML = renderStars(product.Rating);
  dom.pdpRatingNum.textContent = Number(product.Rating).toFixed(1);

  dom.pdpPrice.textContent = price.toLocaleString('en-IN');
  if (changedFields.has('Price (INR)')) dom.pdpPrice.classList.add('value-changed');

  dom.pdpMrp.textContent = `₹${originalPrice.toLocaleString('en-IN')}`;
  dom.pdpDiscount.textContent = `${discountPercent}% OFF`;

  dom.pdpStatusBadge.textContent = product.Status || 'In Stock';
  dom.pdpStatusBadge.className = `pdp-floating-badge status-badge ${statusClass}`;
  if (changedFields.has('Status')) dom.pdpStatusBadge.classList.add('value-changed');

  dom.pdpStockCount.textContent = stock;
  if (changedFields.has('Stock')) dom.pdpStockCount.classList.add('value-changed');

  dom.pdpStatusText.textContent = product.Status || 'In Stock';
  dom.pdpStockFill.style.width = `${stockPercent}%`;
  dom.pdpStockFill.style.backgroundColor = stock > 20 ? '#10b981' : (stock > 0 ? '#f59e0b' : '#f43f5e');

  dom.qtyMaxNote.textContent = isOutOfStock ? 'Item is currently Out of Stock' : `Max available: ${stock} units`;

  // Enable/Disable buy buttons based on real-time stock
  dom.pdpAddCartBtn.disabled = isOutOfStock;
  dom.pdpBuyNowBtn.disabled = isOutOfStock;
  dom.pdpAddCartBtn.innerHTML = isOutOfStock
    ? 'Out of Stock'
    : `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="9" cy="21" r="1"></circle><circle cx="20" cy="21" r="1"></circle><path d="M1 1h4l2.68 13.39a2 2 0 0 0 2 1.61h9.72a2 2 0 0 0 2-1.61L23 6H6"></path></svg> Add to Cart`;
  dom.pdpBuyNowBtn.textContent = isOutOfStock ? 'Unavailable' : 'Buy Now';

  // Reset qty to 1
  dom.pdpQtyInput.value = '1';
}

// ============================================================================
// Shopping Cart Operations (Persistent)
// ============================================================================
function loadCartFromStorage() {
  try {
    const raw = localStorage.getItem('nexus_cart');
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}

function saveCartToStorage() {
  try {
    localStorage.setItem('nexus_cart', JSON.stringify(state.cart));
  } catch (e) {}
}

function addToCart(productId, quantity = 1) {
  const product = state.products.find(p => String(p.ID) === String(productId));
  if (!product) return;

  const maxStock = Number(product.Stock) || 0;
  if (maxStock <= 0) {
    showToast('Out of Stock', `Sorry, ${product.Name} is currently unavailable.`);
    return;
  }

  const existing = state.cart.find(it => String(it.id) === String(productId));
  if (existing) {
    existing.quantity = Math.min(maxStock, existing.quantity + quantity);
  } else {
    state.cart.push({
      id: String(product.ID),
      name: product.Name,
      category: product.Category,
      price: Number(product['Price (INR)']) || 0,
      image: getProductImage(product),
      quantity: Math.min(maxStock, quantity)
    });
  }

  saveCartToStorage();
  updateCartBadge();
  showToast('Added to Cart', `Added <strong>${escapeHtml(product.Name)}</strong> to your shopping cart.`);
}

function removeFromCart(productId) {
  state.cart = state.cart.filter(it => String(it.id) !== String(productId));
  saveCartToStorage();
  updateCartBadge();
  renderCartPage();
  renderDrawerCart();
}

function updateCartQty(productId, newQty) {
  const item = state.cart.find(it => String(it.id) === String(productId));
  if (!item) return;

  const product = state.products.find(p => String(p.ID) === String(productId));
  const maxStock = product ? (Number(product.Stock) || 99) : 99;

  item.quantity = Math.max(1, Math.min(maxStock, newQty));
  saveCartToStorage();
  updateCartBadge();
  renderCartPage();
  renderDrawerCart();
}

function clearCart() {
  state.cart = [];
  saveCartToStorage();
  updateCartBadge();
}

function updateCartBadge() {
  const count = state.cart.reduce((sum, item) => sum + item.quantity, 0);
  dom.navCartBadge.textContent = count;
  dom.drawerItemsCount.textContent = count;
}

function calculateCartTotals() {
  const subtotal = state.cart.reduce((sum, it) => sum + (it.price * it.quantity), 0);
  const discount = Math.round(subtotal * (state.discountPercent / 100));
  const shipping = subtotal > 999 || subtotal === 0 ? 0 : 99;
  const taxableAmount = Math.max(0, subtotal - discount);
  const tax = Math.round(taxableAmount * 0.18); // 18% GST
  const grandTotal = taxableAmount + shipping + tax;

  return { subtotal, discount, shipping, tax, grandTotal };
}

function applyPromoCode() {
  const code = (dom.couponInput.value || '').trim().toUpperCase();
  if (code === 'EXCEL10') {
    state.discountPercent = 10;
    state.activePromoCode = 'EXCEL10';
    dom.couponMessage.textContent = 'Coupon EXCEL10 applied! 10% discount added.';
    dom.couponMessage.style.color = '#34d399';
    dom.discountRow.classList.remove('hidden');
    renderCartPage();
  } else {
    dom.couponMessage.textContent = 'Invalid promo code. Try "EXCEL10".';
    dom.couponMessage.style.color = '#f87171';
  }
}

// ============================================================================
// PAGE 3: Cart Page & Drawer Rendering
// ============================================================================
function renderCartPage() {
  if (state.cart.length === 0) {
    dom.cartContentWrapper.classList.add('hidden');
    dom.cartEmptyView.classList.remove('hidden');
    return;
  }

  dom.cartEmptyView.classList.add('hidden');
  dom.cartContentWrapper.classList.remove('hidden');

  dom.cartItemsList.innerHTML = state.cart.map(item => `
    <div class="cart-item-row" data-id="${escapeHtml(item.id)}">
      <div class="cart-item-info">
        <img src="${item.image}" alt="${escapeHtml(item.name)}" class="cart-item-thumb">
        <div>
          <div class="cart-item-title">${escapeHtml(item.name)}</div>
          <div class="cart-item-cat">${escapeHtml(item.category || 'General')}</div>
          <div class="cart-item-price">₹${Number(item.price).toLocaleString('en-IN')}</div>
        </div>
      </div>
      <div class="cart-item-actions">
        <div class="quantity-stepper">
          <button class="qty-btn" onclick="updateCartQty('${escapeHtml(item.id)}', ${item.quantity - 1})">-</button>
          <input type="text" value="${item.quantity}" readonly>
          <button class="qty-btn" onclick="updateCartQty('${escapeHtml(item.id)}', ${item.quantity + 1})">+</button>
        </div>
        <div style="font-weight: 700; min-width: 80px; text-align: right;">
          ₹${(item.price * item.quantity).toLocaleString('en-IN')}
        </div>
        <button class="btn-remove-item" onclick="removeFromCart('${escapeHtml(item.id)}')" title="Remove item">&times;</button>
      </div>
    </div>
  `).join('');

  const totals = calculateCartTotals();
  dom.cartSubtotal.textContent = `₹${totals.subtotal.toLocaleString('en-IN')}`;
  dom.cartShipping.textContent = totals.shipping === 0 ? 'FREE' : `₹${totals.shipping}`;
  dom.cartTax.textContent = `₹${totals.tax.toLocaleString('en-IN')}`;

  if (totals.discount > 0) {
    dom.discountRow.classList.remove('hidden');
    dom.cartDiscount.textContent = `-₹${totals.discount.toLocaleString('en-IN')}`;
  } else {
    dom.discountRow.classList.add('hidden');
  }

  dom.cartGrandTotal.textContent = `₹${totals.grandTotal.toLocaleString('en-IN')}`;
}

function openCartDrawer() {
  renderDrawerCart();
  dom.drawerOverlay.classList.remove('hidden');
}

function closeCartDrawer() {
  dom.drawerOverlay.classList.add('hidden');
}

function renderDrawerCart() {
  updateCartBadge();
  const totals = calculateCartTotals();
  dom.drawerSubtotal.textContent = `₹${totals.subtotal.toLocaleString('en-IN')}`;

  if (state.cart.length === 0) {
    dom.drawerItemsList.innerHTML = `<div style="text-align: center; color: #94a3b8; padding: 2rem 0;">Your cart is empty.</div>`;
    dom.drawerCheckoutBtn.classList.add('hidden');
    return;
  }

  dom.drawerCheckoutBtn.classList.remove('hidden');
  dom.drawerItemsList.innerHTML = state.cart.map(it => `
    <div class="drawer-item">
      <img src="${it.image}" alt="${escapeHtml(it.name)}" class="drawer-item-img">
      <div class="drawer-item-info">
        <div class="drawer-item-title">${escapeHtml(it.name)}</div>
        <div class="drawer-item-price">Qty: ${it.quantity} &times; ₹${it.price.toLocaleString('en-IN')}</div>
      </div>
      <button class="btn-remove-item" onclick="removeFromCart('${escapeHtml(it.id)}')">&times;</button>
    </div>
  `).join('');
}

// ============================================================================
// PAGE 4: Checkout Page
// ============================================================================
function renderCheckoutPage() {
  if (state.cart.length === 0) {
    navigateTo('#cart');
    return;
  }

  dom.checkoutItemsPreview.innerHTML = state.cart.map(it => `
    <div class="checkout-item-mini">
      <span>${escapeHtml(it.name)} &times; ${it.quantity}</span>
      <strong>₹${(it.price * it.quantity).toLocaleString('en-IN')}</strong>
    </div>
  `).join('');

  const totals = calculateCartTotals();
  dom.checkoutSubtotal.textContent = `₹${totals.subtotal.toLocaleString('en-IN')}`;
  dom.checkoutTax.textContent = `₹${totals.tax.toLocaleString('en-IN')}`;

  if (totals.discount > 0) {
    dom.checkoutDiscountRow.classList.remove('hidden');
    dom.checkoutDiscount.textContent = `-₹${totals.discount.toLocaleString('en-IN')}`;
  } else {
    dom.checkoutDiscountRow.classList.add('hidden');
  }

  dom.checkoutGrandTotal.textContent = `₹${totals.grandTotal.toLocaleString('en-IN')}`;
}

async function handlePlaceOrder(e) {
  e.preventDefault();

  if (state.cart.length === 0) {
    alert('Cart is empty.');
    return;
  }

  const selectedPayment = document.querySelector('input[name="payment-method"]:checked').value;
  const customer = {
    fullName: document.getElementById('cust-name').value.trim(),
    email: document.getElementById('cust-email').value.trim(),
    phone: document.getElementById('cust-phone').value.trim(),
    address: document.getElementById('cust-address').value.trim(),
    city: document.getElementById('cust-city').value.trim(),
    pincode: document.getElementById('cust-pincode').value.trim()
  };

  const totals = calculateCartTotals();

  // If Razorpay is selected, launch the authentic Razorpay Checkout modal
  if (selectedPayment === 'Razorpay') {
    openRazorpayModal(customer, totals);
    return;
  }

  // Else Cash on Delivery flow
  dom.placeOrderBtn.disabled = true;
  dom.placeOrderBtn.innerHTML = `Processing Order...`;

  try {
    const res = await fetch('/api/orders', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items: state.cart,
        customer,
        paymentMethod: 'Cash on Delivery (COD)',
        totals
      })
    });

    const data = await res.json();
    if (!res.ok || !data.success) throw new Error(data.error || 'Failed to place order');

    state.lastOrder = data.order;
    clearCart();
    navigateTo(`#order-success?orderId=${data.order.orderId}`);
  } catch (err) {
    alert('Error placing order: ' + err.message);
  } finally {
    dom.placeOrderBtn.disabled = false;
    dom.placeOrderBtn.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg> Confirm & Place Order`;
  }
}

// ============================================================================
// RAZORPAY CHECKOUT MODAL LOGIC (AUTHENTIC DEMO)
// ============================================================================
state.pendingRazorpayOrder = null;
state.rzpTimerId = null;

function setupRazorpayModalListeners() {
  // Tab Switching
  document.querySelectorAll('.rzp-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.rzp-tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.rzp-tab-panel').forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      const targetPanel = document.getElementById(`rzp-panel-${tab.dataset.tab}`);
      if (targetPanel) targetPanel.classList.add('active');
    });
  });

  // Netbanking bank cards
  document.querySelectorAll('.rzp-bank-card').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelectorAll('.rzp-bank-card').forEach(c => c.classList.remove('active'));
      card.classList.add('active');
    });
  });

  // Wallets selector
  document.querySelectorAll('.rzp-wallet-item').forEach(item => {
    item.addEventListener('click', () => {
      document.querySelectorAll('.rzp-wallet-item').forEach(w => w.classList.remove('active'));
      item.classList.add('active');
      const radio = item.querySelector('input');
      if (radio) radio.checked = true;
    });
  });

  // Close Razorpay Modal
  dom.rzpCloseBtn.addEventListener('click', closeRazorpayModal);

  // Trigger simulated payment success from any method:
  dom.rzpSimulateSuccessBtn.addEventListener('click', () => executeRazorpayPayment(true));
  dom.rzpCardPayBtn.addEventListener('click', () => executeRazorpayPayment(true));
  dom.rzpUpiVerifyPayBtn.addEventListener('click', () => executeRazorpayPayment(true));
  dom.rzpNetbankingPayBtn.addEventListener('click', () => executeRazorpayPayment(true));
  dom.rzpWalletPayBtn.addEventListener('click', () => executeRazorpayPayment(true));

  // Trigger simulated payment decline:
  dom.rzpSimulateFailureBtn.addEventListener('click', () => executeRazorpayPayment(false));
}

function openRazorpayModal(customer, totals) {
  state.pendingRazorpayOrder = { customer, totals };
  const tempOrderId = 'ORD-' + Math.floor(100000 + Math.random() * 900000);
  const amountFormatted = `₹${totals.grandTotal.toLocaleString('en-IN')}`;

  dom.rzpOrderPreviewId.textContent = tempOrderId;
  dom.rzpAmountDisplay.textContent = amountFormatted;
  document.querySelectorAll('.rzp-btn-amount').forEach(el => el.textContent = amountFormatted);

  dom.rzpProcessingOverlay.classList.add('hidden');
  dom.rzpModal.classList.remove('hidden');

  // Start 5-minute QR code timer
  startQrCountdown();
}

function closeRazorpayModal() {
  if (confirm('Cancel Razorpay payment? Your cart items will remain saved.')) {
    dom.rzpModal.classList.add('hidden');
    dom.rzpProcessingOverlay.classList.add('hidden');
    if (state.rzpTimerId) clearInterval(state.rzpTimerId);
  }
}

function startQrCountdown() {
  if (state.rzpTimerId) clearInterval(state.rzpTimerId);
  let secondsRemaining = 299; // 04:59

  const updateDisplay = () => {
    const mins = Math.floor(secondsRemaining / 60);
    const secs = secondsRemaining % 60;
    dom.rzpQrTimerText.textContent = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  };

  updateDisplay();
  state.rzpTimerId = setInterval(() => {
    secondsRemaining--;
    if (secondsRemaining <= 0) {
      clearInterval(state.rzpTimerId);
      dom.rzpQrTimerText.textContent = 'Expired';
    } else {
      updateDisplay();
    }
  }, 1000);
}

async function executeRazorpayPayment(isSuccess) {
  dom.rzpProcessingOverlay.classList.remove('hidden');

  if (!isSuccess) {
    setTimeout(() => {
      dom.rzpProcessingOverlay.classList.add('hidden');
      showToast('Payment Declined', 'Sandbox Error: 1002 - Bank authorization declined by cardholder/bank. Please try again.');
    }, 900);
    return;
  }

  // Simulate bank processing time (1.2 seconds)
  setTimeout(async () => {
    const randomPayId = 'pay_' + Math.random().toString(36).substring(2, 8).toUpperCase() + Math.random().toString(36).substring(2, 7).toUpperCase();
    const { customer, totals } = state.pendingRazorpayOrder;

    try {
      const res = await fetch('/api/orders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: state.cart,
          customer,
          paymentMethod: 'Razorpay Secure Gateway',
          razorpayPaymentId: randomPayId,
          totals
        })
      });

      const data = await res.json();
      if (!res.ok || !data.success) throw new Error(data.error || 'Failed to complete order');

      state.lastOrder = data.order;
      clearCart();
      dom.rzpModal.classList.add('hidden');
      dom.rzpProcessingOverlay.classList.add('hidden');
      if (state.rzpTimerId) clearInterval(state.rzpTimerId);

      navigateTo(`#order-success?orderId=${data.order.orderId}`);
      showToast('Payment Verified', `Razorpay Payment <strong>${randomPayId}</strong> captured successfully!`);
    } catch (err) {
      dom.rzpProcessingOverlay.classList.add('hidden');
      alert('Order placement error: ' + err.message);
    }
  }, 1200);
}

// ============================================================================
// PAGE 5: Order Success Page
// ============================================================================
function renderOrderSuccessPage() {
  const order = state.lastOrder;
  if (!order) {
    navigateTo('#catalog');
    return;
  }

  dom.successOrderId.textContent = order.orderId;
  dom.successDeliveryDate.textContent = order.estimatedDelivery;
  dom.successPaymentMethod.textContent = order.paymentMethod;
  dom.successCustomerName.textContent = `${order.customer.fullName} (${order.customer.city} - ${order.customer.pincode})`;
  dom.successTotalPaid.textContent = `₹${order.totals.grandTotal.toLocaleString('en-IN')}`;
  dom.successTimestamp.textContent = new Date(order.createdAt).toLocaleString('en-IN');

  // Handle Razorpay Payment ID row
  if (order.razorpayPaymentId) {
    dom.successRazorpayIdRow.classList.remove('hidden');
    dom.successRazorpayId.textContent = order.razorpayPaymentId;
    dom.successPaymentStatus.textContent = 'Paid & Captured (Verified by Razorpay)';
    dom.successPaymentStatus.className = 'status-badge status-in-stock';
  } else {
    dom.successRazorpayIdRow.classList.add('hidden');
    dom.successPaymentStatus.textContent = 'Pending (Cash on Delivery)';
    dom.successPaymentStatus.className = 'status-badge status-low-stock';
  }

  dom.successItemsList.innerHTML = order.items.map(it => `
    <div class="receipt-item-line">
      <span>${escapeHtml(it.name)} &times; ${it.quantity}</span>
      <strong>₹${(it.price * it.quantity).toLocaleString('en-IN')}</strong>
    </div>
  `).join('');
}

// ============================================================================
// PAGE 6: Excel Live Monitor / Admin Diagnostics
// ============================================================================
async function renderAdminPage() {
  try {
    const res = await fetch('/api/meta');
    const meta = await res.json();
    dom.adminRowCount.textContent = `${meta.rowCount || state.products.length} products`;
    if (meta.lastModified) {
      dom.adminLastModified.textContent = new Date(meta.lastModified).toLocaleTimeString('en-IN', {
        hour: '2-digit', minute: '2-digit', second: '2-digit'
      });
    }
  } catch (e) {
    dom.adminRowCount.textContent = `${state.products.length} products`;
  }
}

// ============================================================================
// Utilities
// ============================================================================
function getStatusClass(status) {
  const s = String(status || '').trim().toLowerCase();
  if (s === 'in stock') return 'status-in-stock';
  if (s === 'low stock') return 'status-low-stock';
  if (s === 'out of stock') return 'status-out-stock';
  return 'status-in-stock';
}

function renderStars(rating) {
  const r = Math.round(Number(rating) || 0);
  let str = '';
  for (let i = 1; i <= 5; i++) str += i <= r ? '★' : '☆';
  return str;
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
