import React, { useState } from 'react';
import * as XLSX from 'xlsx';

function App() {
  const [view, setView] = useState('customer'); // 'customer', 'customerLogin', 'adminLogin', 'adminPanel'
  const [adminPassword, setAdminPassword] = useState('');
  const [adminError, setAdminError] = useState('');
  
  // Active Tab for Admin Sidebar
  const [adminTab, setAdminTab] = useState('Dashboard');

  // Customer Login State
  const [customerPhone, setCustomerPhone] = useState('');
  const [isCustomerLoggedIn, setIsCustomerLoggedIn] = useState(false);

  // Store Settings State
  const [storeSettings, setStoreSettings] = useState({
    storeName: 'One Click Online Services',
    phone: '9876543210',
    address: 'Kinwat, Maharashtra',
    deliveryFee: 15,
    minFreeDelivery: 199
  });

  // Products State
  const [products, setProducts] = useState([
    { id: 1, name: 'Amul Taaza Milk', price: 27, mrp: 30, unit: '500 ml', stock: 50, category: 'Dairy' },
    { id: 2, name: 'Brown Bread', price: 40, mrp: 45, unit: '400 g', stock: 20, category: 'Bakery' }
  ]);

  // Categories State
  const [categories, setCategories] = useState(['Dairy', 'Bakery', 'Groceries', 'Snacks', 'Beverages']);
  const [newCategoryName, setNewCategoryName] = useState('');

  // REAL Dynamic Orders State (Empty initially, fills when customers order)
  const [orders, setOrders] = useState([]);

  const [searchTerm, setSearchTerm] = useState('');
  const [newProduct, setNewProduct] = useState({ name: '', price: '', mrp: '', unit: '', stock: '', category: 'Dairy' });

  // ------------------------------------------------------------------
  // CART & CHECKOUT STATE
  // ------------------------------------------------------------------
  const [cart, setCart] = useState([]);
  const [isCartOpen, setIsCartOpen] = useState(false);
  const [checkoutStep, setCheckoutStep] = useState('cart'); // 'cart', 'address', 'payment', 'summary', 'confirmed'
  
  // Delivery Address State
  const [address, setAddress] = useState({
    fullName: 'Avinash Rathod',
    phone: '9876543210',
    houseNo: 'Shop No. 1, Main Market',
    area: 'Station Road',
    city: 'Kinwat',
    pincode: '431804'
  });
  const [isEditingAddress, setIsEditingAddress] = useState(false);

  // Payment Method State
  const [paymentMethod, setPaymentMethod] = useState('upi');

  // Cart Functions
  const addToCart = (product) => {
    setCart(prevCart => {
      const existing = prevCart.find(item => item.id === product.id);
      if (existing) {
        return prevCart.map(item =>
          item.id === product.id ? { ...item, qty: item.qty + 1 } : item
        );
      }
      return [...prevCart, { ...product, qty: 1 }];
    });
    setIsCartOpen(true);
  };

  const updateQuantity = (id, delta) => {
    setCart(prevCart =>
      prevCart
        .map(item => {
          if (item.id === id) {
            const newQty = item.qty + delta;
            return newQty > 0 ? { ...item, qty: newQty } : null;
          }
          return item;
        })
        .filter(Boolean)
    );
  };

  // Calculations
  const itemsTotal = cart.reduce((sum, item) => sum + item.price * item.qty, 0);
  const mrpTotal = cart.reduce((sum, item) => sum + (item.mrp || item.price) * item.qty, 0);
  const discount = mrpTotal - itemsTotal;
  const deliveryFee = itemsTotal > storeSettings.minFreeDelivery || itemsTotal === 0 ? 0 : storeSettings.deliveryFee;
  const platformFee = itemsTotal > 0 ? 2 : 0;
  const grandTotal = itemsTotal + deliveryFee + platformFee;

  // Real Dynamic Order Placement
  const handlePlaceOrder = () => {
    const newOrder = {
      id: '#OCM' + Math.floor(100000 + Math.random() * 900000),
      customerName: address.fullName,
      phone: address.phone || customerPhone || '9876543210',
      address: `${address.houseNo}, ${address.area}, ${address.city}`,
      items: cart,
      totalAmount: grandTotal,
      paymentMethod: paymentMethod.toUpperCase(),
      status: 'Placed',
      date: new Date().toLocaleString()
    };

    // Update orders list
    setOrders(prev => [newOrder, ...prev]);

    // Decrease Product Stock
    setProducts(prevProducts =>
      prevProducts.map(p => {
        const cartItem = cart.find(ci => ci.id === p.id);
        if (cartItem) {
          return { ...p, stock: Math.max(0, p.stock - cartItem.qty) };
        }
        return p;
      })
    );

    setCheckoutStep('confirmed');
  };

  // Handlers for Admin/Customer Auth
  const handleAdminLogin = (e) => {
    e.preventDefault();
    if (adminPassword === 'admin') {
      setView('adminPanel');
      setAdminError('');
      setAdminPassword('');
    } else {
      setAdminError('गलत पासवर्ड! कृपया पुन्हा प्रयत्न करा.');
    }
  };

  const handleCustomerLogin = (e) => {
    e.preventDefault();
    if (customerPhone.length === 10) {
      setIsCustomerLoggedIn(true);
      setView('customer');
    } else {
      alert('कृपया १० अंकी मोबाईल नंबर टाका');
    }
  };

  const handleAddProduct = (e) => {
    e.preventDefault();
    setProducts([...products, { ...newProduct, id: Date.now(), price: Number(newProduct.price), mrp: Number(newProduct.mrp), stock: Number(newProduct.stock) }]);
    setNewProduct({ name: '', price: '', mrp: '', unit: '', stock: '', category: categories[0] || 'General' });
  };

  const handleAddCategory = (e) => {
    e.preventDefault();
    if (newCategoryName.trim() && !categories.includes(newCategoryName.trim())) {
      setCategories([...categories, newCategoryName.trim()]);
      setNewCategoryName('');
    }
  };

  const handleDelete = (id) => {
    setProducts(products.filter(p => p.id !== id));
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      
      const formattedData = data.map((item, index) => ({
        id: Date.now() + index,
        name: item.Name || item.name || 'Unnamed Product',
        price: Number(item.Price || item.price || 0),
        mrp: Number(item.MRP || item.mrp || 0),
        unit: item.Unit || item.unit || '1 unit',
        stock: Number(item.Stock || item.stock || 0),
        category: item.Category || item.category || 'General'
      }));

      setProducts(prev => [...prev, ...formattedData]);
      alert(`${formattedData.length} प्रॉडक्ट्स यशस्वीरीत्या अपलोड झाले!`);
    };
    reader.readAsBinaryString(file);
  };

  const filteredProducts = products.filter(item =>
    item.name.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const resetCheckoutFlow = () => {
    setIsCartOpen(false);
    setCheckoutStep('cart');
  };

  // Dynamic Dashboard Stats Calculations
  const totalSalesAmount = orders.reduce((sum, order) => sum + order.totalAmount, 0);
  const uniqueCustomersCount = new Set(orders.map(o => o.phone)).size;
  const lowStockProducts = products.filter(p => p.stock <= 10);

  // Group unique customers for Customers Tab
  const customersList = Object.values(orders.reduce((acc, order) => {
    if (!acc[order.phone]) {
      acc[order.phone] = { name: order.customerName, phone: order.phone, totalOrders: 0, totalSpent: 0 };
    }
    acc[order.phone].totalOrders += 1;
    acc[order.phone].totalSpent += order.totalAmount;
    return acc;
  }, {}));

  return (
    <div style={{ fontFamily: 'sans-serif', backgroundColor: '#f8f9fa', minHeight: '100vh', margin: 0 }}>
      
      {/* HEADER (Only shown outside Admin Panel) */}
      {view !== 'adminPanel' && (
        <header style={{ 
          backgroundColor: '#f7c223', 
          padding: '10px 30px', 
          display: 'flex', 
          justifyContent: 'space-between', 
          alignItems: 'center',
          boxShadow: '0 2px 4px rgba(0,0,0,0.08)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
            <h1 style={{ margin: 0, fontSize: '26px', color: '#0c831f', fontWeight: '800', cursor: 'pointer' }} onClick={() => setView('customer')}>
              One Click
            </h1>
            <div style={{ borderLeft: '1px solid #d1a51b', paddingLeft: '12px', fontSize: '11px', color: '#333' }}>
              <div style={{ fontWeight: 'bold' }}>Delivery in 10 minutes</div>
              <div style={{ color: '#555' }}>{storeSettings.address}</div>
            </div>
          </div>

          {/* Search Bar */}
          <div style={{ position: 'relative', width: '40%' }}>
            <input 
              type="text" 
              placeholder="Search products..." 
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={{ width: '100%', padding: '9px 15px 9px 35px', borderRadius: '8px', border: 'none', outline: 'none', fontSize: '14px' }}
            />
            <span style={{ position: 'absolute', left: '12px', top: '8px', color: '#888' }}>🔍</span>
          </div>

          {/* Right Buttons */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            {isCustomerLoggedIn ? (
              <span style={{ fontWeight: 'bold', fontSize: '14px', color: '#0c831f' }}>👤 Profile</span>
            ) : (
              <button onClick={() => setView('customerLogin')} style={{ background: 'none', border: 'none', fontWeight: 'bold', cursor: 'pointer', fontSize: '14px', color: '#333' }}>
                Login
              </button>
            )}

            <button onClick={() => setView(view === 'adminPanel' ? 'customer' : 'adminLogin')} style={{ backgroundColor: '#000', color: '#fff', border: 'none', padding: '7px 12px', borderRadius: '6px', cursor: 'pointer', fontWeight: 'bold', fontSize: '12px' }}>
              {view === 'adminPanel' ? '🛒 Customer View' : '🔒 Admin'}
            </button>

            <button onClick={() => setIsCartOpen(true)} style={{ backgroundColor: '#0c831f', color: 'white', border: 'none', padding: '8px 14px', borderRadius: '6px', fontWeight: 'bold', display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '13px' }}>
              🛒 {cart.reduce((a, b) => a + b.qty, 0)} Items
            </button>
          </div>
        </header>
      )}

      {/* VIEW 1: STORE FRONT */}
      {view === 'customer' && (
        <div>
          <div style={{ backgroundColor: '#0c831f', color: 'white', textAlign: 'center', padding: '15px' }}>
            <h2 style={{ margin: 0, fontSize: '20px', fontWeight: '700' }}>Fresh Groceries Delivered at Your Doorstep!</h2>
            <p style={{ margin: '4px 0 0 0', fontSize: '13px', opacity: 0.9 }}>Best quality products at low prices in Kinwat.</p>
          </div>

          <main style={{ padding: '30px 40px' }}>
            <h3 style={{ fontSize: '18px', color: '#222', marginBottom: '20px' }}>Popular Products ({filteredProducts.length})</h3>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '20px' }}>
              {filteredProducts.map(p => (
                <div key={p.id} style={{ background: '#fff', padding: '15px', borderRadius: '10px', boxShadow: '0 2px 8px rgba(0,0,0,0.06)', textAlign: 'center' }}>
                  <h4 style={{ margin: '5px 0' }}>{p.name}</h4>
                  <p style={{ color: '#666', fontSize: '12px', margin: '0 0 8px' }}>{p.unit} | Stock: {p.stock}</p>
                  <p style={{ margin: '0 0 10px' }}>
                    <strong style={{ fontSize: '16px' }}>₹{p.price}</strong> 
                    <span style={{ textDecoration: 'line-through', color: '#888', fontSize: '12px', marginLeft: '6px' }}>₹{p.mrp}</span>
                  </p>
                  <button 
                    disabled={p.stock <= 0}
                    onClick={() => addToCart(p)} 
                    style={{ backgroundColor: p.stock > 0 ? '#0c831f' : '#ccc', color: 'white', border: 'none', width: '100%', padding: '7px', borderRadius: '5px', fontWeight: 'bold', cursor: p.stock > 0 ? 'pointer' : 'not-allowed' }}
                  >
                    {p.stock > 0 ? 'ADD' : 'OUT OF STOCK'}
                  </button>
                </div>
              ))}
            </div>
          </main>
        </div>
      )}

      {/* VIEW 2: CUSTOMER LOGIN */}
      {view === 'customerLogin' && (
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: '65vh' }}>
          <form onSubmit={handleCustomerLogin} style={{ background: '#fff', padding: '30px', borderRadius: '10px', boxShadow: '0 4px 10px rgba(0,0,0,0.08)', width: '320px', textAlign: 'center' }}>
            <h2 style={{ color: '#0c831f', marginBottom: '10px' }}>Customer Login</h2>
            <p style={{ fontSize: '13px', color: '#666', marginBottom: '20px' }}>Enter mobile number to order groceries</p>
            <input type="tel" placeholder="Mobile Number (10 digits)" value={customerPhone} onChange={(e) => setCustomerPhone(e.target.value)} style={{ width: '100%', padding: '10px', marginBottom: '15px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
            <button type="submit" style={{ backgroundColor: '#0c831f', color: 'white', border: 'none', width: '100%', padding: '10px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>Continue</button>
          </form>
        </div>
      )}

      {/* VIEW 3: ADMIN LOGIN */}
      {view === 'adminLogin' && (
        <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: '65vh' }}>
          <form onSubmit={handleAdminLogin} style={{ background: '#fff', padding: '30px', borderRadius: '10px', boxShadow: '0 4px 10px rgba(0,0,0,0.08)', width: '320px', textAlign: 'center' }}>
            <h2 style={{ color: '#000', marginBottom: '10px' }}>🔒 Admin Portal</h2>
            <p style={{ fontSize: '13px', color: '#666', marginBottom: '20px' }}>Restricted access for store management</p>
            {adminError && <p style={{ color: 'red', fontSize: '13px' }}>{adminError}</p>}
            <input type="password" placeholder="Enter Admin Password" value={adminPassword} onChange={(e) => setAdminPassword(e.target.value)} style={{ width: '100%', padding: '10px', marginBottom: '15px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' }} required />
            <button type="submit" style={{ backgroundColor: '#000', color: 'white', border: 'none', width: '100%', padding: '10px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>Login to Dashboard</button>
            <p style={{ fontSize: '12px', color: '#888', marginTop: '12px' }}>Default Password: <strong>admin</strong></p>
          </form>
        </div>
      )}

      {/* VIEW 4: ADMIN PANEL */}
      {view === 'adminPanel' && (
        <div style={{ display: 'flex', minHeight: '100vh', backgroundColor: '#f4f6f8' }}>
          
          {/* LEFT SIDEBAR */}
          <div style={{ width: '240px', backgroundColor: '#0d1b2a', color: '#fff', padding: '20px 0', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
            <div>
              <div style={{ padding: '0 20px 20px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ backgroundColor: '#f7c223', padding: '8px', borderRadius: '8px', color: '#000', fontWeight: 'bold', fontSize: '18px' }}>🛒</div>
                <div>
                  <div style={{ fontWeight: 'bold', fontSize: '16px', color: '#f7c223' }}>One Click</div>
                  <div style={{ fontSize: '12px', color: '#aaa' }}>Multiservice</div>
                  <div style={{ fontSize: '11px', color: '#666', marginTop: '2px' }}>Admin Panel</div>
                </div>
              </div>

              {/* Navigation Links */}
              <nav style={{ display: 'flex', flexDirection: 'column', gap: '5px', padding: '0 10px' }}>
                {[
                  { name: 'Dashboard', icon: '🏠' },
                  { name: 'Products', icon: '📦' },
                  { name: 'Categories', icon: '📑' },
                  { name: 'Orders', icon: '🛒' },
                  { name: 'Customers', icon: '👥' },
                  { name: 'Inventory', icon: '🏭' },
                  { name: 'Sales Report', icon: '📊' },
                  { name: 'Settings', icon: '⚙️' }
                ].map(item => {
                  const isActive = adminTab === item.name;
                  return (
                    <button
                      key={item.name}
                      onClick={() => setAdminTab(item.name)}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '12px',
                        padding: '10px 15px',
                        borderRadius: '6px',
                        border: 'none',
                        backgroundColor: isActive ? '#0c831f' : 'transparent',
                        color: isActive ? '#fff' : '#a0aec0',
                        fontWeight: isActive ? 'bold' : 'normal',
                        cursor: 'pointer',
                        textAlign: 'left',
                        fontSize: '14px'
                      }}
                    >
                      <span>{item.icon}</span>
                      <span>{item.name}</span>
                    </button>
                  );
                })}
              </nav>
            </div>

            {/* Logout Button */}
            <div style={{ padding: '0 10px' }}>
              <button
                onClick={() => setView('customer')}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '12px',
                  padding: '10px 15px',
                  width: '100%',
                  borderRadius: '6px',
                  border: 'none',
                  backgroundColor: 'transparent',
                  color: '#a0aec0',
                  cursor: 'pointer',
                  fontSize: '14px'
                }}
              >
                <span>🚪</span>
                <span>Logout</span>
              </button>
            </div>
          </div>

          {/* MAIN CONTENT AREA */}
          <div style={{ flex: 1, padding: '25px 30px', overflowY: 'auto' }}>
            
            {/* 1. DASHBOARD VIEW (FULLY DYNAMIC DATA) */}
            {adminTab === 'Dashboard' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Dashboard</h2>

                {/* Dynamic Metrics Cards */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '20px', marginBottom: '25px' }}>
                  
                  {/* Total Orders */}
                  <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '15px' }}>
                    <div style={{ backgroundColor: '#e6f4ea', color: '#0c831f', padding: '12px', borderRadius: '8px', fontSize: '20px' }}>🛒</div>
                    <div>
                      <div style={{ fontSize: '12px', color: '#718096' }}>Total Orders</div>
                      <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#1a202c', margin: '2px 0' }}>{orders.length}</div>
                      <div style={{ fontSize: '12px', color: '#38a169', fontWeight: 'bold' }}>Live Real Data</div>
                    </div>
                  </div>

                  {/* Total Customers */}
                  <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '15px' }}>
                    <div style={{ backgroundColor: '#ebf8ff', color: '#3182ce', padding: '12px', borderRadius: '8px', fontSize: '20px' }}>👥</div>
                    <div>
                      <div style={{ fontSize: '12px', color: '#718096' }}>Total Customers</div>
                      <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#1a202c', margin: '2px 0' }}>{uniqueCustomersCount}</div>
                      <div style={{ fontSize: '12px', color: '#3182ce', fontWeight: 'bold' }}>Unique Buyers</div>
                    </div>
                  </div>

                  {/* Total Sales */}
                  <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', alignItems: 'center', gap: '15px' }}>
                    <div style={{ backgroundColor: '#faf5ff', color: '#805ad5', padding: '12px', borderRadius: '8px', fontSize: '20px' }}>💰</div>
                    <div>
                      <div style={{ fontSize: '12px', color: '#718096' }}>Total Sales</div>
                      <div style={{ fontSize: '22px', fontWeight: 'bold', color: '#1a202c', margin: '2px 0' }}>₹{totalSalesAmount}</div>
                      <div style={{ fontSize: '12px', color: '#805ad5', fontWeight: 'bold' }}>Total Revenue</div>
                    </div>
                  </div>

                </div>

                {/* Recent Orders Table */}
                <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', marginBottom: '25px' }}>
                  <h3 style={{ margin: '0 0 15px', fontSize: '16px', color: '#2d3748' }}>Recent Orders ({orders.length})</h3>
                  {orders.length === 0 ? (
                    <p style={{ color: '#718096', fontSize: '14px', margin: 0 }}>Aapke customers ke dwara koi order place hote hi yaha real details dikhai dengi.</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ color: '#a0aec0', borderBottom: '1px solid #edf2f7', textAlign: 'left' }}>
                          <th style={{ paddingBottom: '10px' }}>Order ID</th>
                          <th style={{ paddingBottom: '10px' }}>Customer</th>
                          <th style={{ paddingBottom: '10px' }}>Phone</th>
                          <th style={{ paddingBottom: '10px' }}>Amount</th>
                          <th style={{ paddingBottom: '10px' }}>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {orders.slice(0, 5).map((order) => (
                          <tr key={order.id} style={{ borderBottom: '1px solid #f7fafc' }}>
                            <td style={{ padding: '10px 0', color: '#718096', fontWeight: 'bold' }}>{order.id}</td>
                            <td style={{ fontWeight: '500' }}>{order.customerName}</td>
                            <td>{order.phone}</td>
                            <td>₹{order.totalAmount}</td>
                            <td>
                              <span style={{ backgroundColor: '#c6f6d5', color: '#22543d', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
                                {order.status}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>

                {/* Low Stock Alert */}
                <div style={{ backgroundColor: '#fff', padding: '20px', borderRadius: '10px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <h3 style={{ margin: '0 0 15px', fontSize: '16px', color: '#2d3748' }}>Low Stock Alert ({lowStockProducts.length})</h3>
                  {lowStockProducts.length === 0 ? (
                    <p style={{ color: '#38a169', fontSize: '13px', margin: 0 }}>Sabhi products ka stock sufficient hai!</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ color: '#a0aec0', borderBottom: '1px solid #edf2f7', textAlign: 'left' }}>
                          <th style={{ paddingBottom: '8px' }}>Product</th>
                          <th style={{ paddingBottom: '8px' }}>Category</th>
                          <th style={{ paddingBottom: '8px', textAlign: 'right' }}>Current Stock</th>
                        </tr>
                      </thead>
                      <tbody>
                        {lowStockProducts.map((p) => (
                          <tr key={p.id} style={{ borderBottom: '1px solid #f7fafc' }}>
                            <td style={{ padding: '8px 0', fontWeight: '500' }}>{p.name}</td>
                            <td style={{ color: '#718096' }}>{p.category}</td>
                            <td style={{ textAlign: 'right', color: '#e53e3e', fontWeight: 'bold' }}>{p.stock}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>

              </div>
            )}

            {/* 2. PRODUCTS SECTION */}
            {adminTab === 'Products' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Manage Products ({products.length})</h2>

                <div style={{ background: '#e9f5ed', border: '2px dashed #0c831f', padding: '20px', borderRadius: '8px', marginBottom: '25px', textAlign: 'center' }}>
                  <h3 style={{ margin: '0 0 10px 0', color: '#0c831f' }}>📥 Bulk Import Excel / CSV (30,000+ Products)</h3>
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} />
                </div>

                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '20px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <h3 style={{ marginTop: 0 }}>➕ Add New Product</h3>
                  <form onSubmit={handleAddProduct} style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                    <input type="text" placeholder="Name" value={newProduct.name} onChange={(e) => setNewProduct({...newProduct, name: e.target.value})} required style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                    <input type="number" placeholder="Price (₹)" value={newProduct.price} onChange={(e) => setNewProduct({...newProduct, price: e.target.value})} required style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                    <input type="number" placeholder="MRP (₹)" value={newProduct.mrp} onChange={(e) => setNewProduct({...newProduct, mrp: e.target.value})} style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                    <input type="text" placeholder="Unit (500ml)" value={newProduct.unit} onChange={(e) => setNewProduct({...newProduct, unit: e.target.value})} required style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                    <input type="number" placeholder="Stock" value={newProduct.stock} onChange={(e) => setNewProduct({...newProduct, stock: e.target.value})} required style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                    <select value={newProduct.category} onChange={(e) => setNewProduct({...newProduct, category: e.target.value})} style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }}>
                      {categories.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                    <button type="submit" style={{ background: '#0c831f', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '5px', cursor: 'pointer', fontWeight: 'bold' }}>Add Item</button>
                  </form>
                </div>

                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: '10px' }}>
                    <thead>
                      <tr style={{ background: '#f2f2f2', textAlign: 'left' }}>
                        <th style={{ padding: '10px' }}>Name</th>
                        <th style={{ padding: '10px' }}>Category</th>
                        <th style={{ padding: '10px' }}>Price</th>
                        <th style={{ padding: '10px' }}>MRP</th>
                        <th style={{ padding: '10px' }}>Unit</th>
                        <th style={{ padding: '10px' }}>Stock</th>
                        <th style={{ padding: '10px' }}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {products.map(p => (
                        <tr key={p.id} style={{ borderBottom: '1px solid #eee' }}>
                          <td style={{ padding: '10px', fontWeight: 'bold' }}>{p.name}</td>
                          <td style={{ padding: '10px' }}>{p.category}</td>
                          <td style={{ padding: '10px' }}>₹{p.price}</td>
                          <td style={{ padding: '10px' }}>₹{p.mrp}</td>
                          <td style={{ padding: '10px' }}>{p.unit}</td>
                          <td style={{ padding: '10px' }}>{p.stock}</td>
                          <td style={{ padding: '10px' }}>
                            <button onClick={() => handleDelete(p.id)} style={{ background: '#dc3545', color: 'white', border: 'none', padding: '4px 10px', borderRadius: '4px', cursor: 'pointer' }}>Delete</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* 3. CATEGORIES SECTION */}
            {adminTab === 'Categories' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Categories ({categories.length})</h2>
                
                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', marginBottom: '20px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <h3 style={{ marginTop: 0 }}>➕ Add New Category</h3>
                  <form onSubmit={handleAddCategory} style={{ display: 'flex', gap: '10px' }}>
                    <input type="text" placeholder="Category Name" value={newCategoryName} onChange={(e) => setNewCategoryName(e.target.value)} required style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px', width: '250px' }} />
                    <button type="submit" style={{ background: '#0c831f', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: '5px', cursor: 'pointer', fontWeight: 'bold' }}>Add Category</button>
                  </form>
                </div>

                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: '#f2f2f2', textAlign: 'left' }}>
                        <th style={{ padding: '10px' }}>Category Name</th>
                        <th style={{ padding: '10px' }}>Products Count</th>
                        <th style={{ padding: '10px' }}>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {categories.map((cat, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid #eee' }}>
                          <td style={{ padding: '10px', fontWeight: 'bold' }}>{cat}</td>
                          <td style={{ padding: '10px' }}>{products.filter(p => p.category === cat).length} items</td>
                          <td style={{ padding: '10px' }}>
                            <button onClick={() => setCategories(categories.filter(c => c !== cat))} style={{ background: '#dc3545', color: 'white', border: 'none', padding: '4px 10px', borderRadius: '4px', cursor: 'pointer' }}>Remove</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* 4. ORDERS SECTION */}
            {adminTab === 'Orders' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Customer Orders ({orders.length})</h2>
                
                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  {orders.length === 0 ? (
                    <p style={{ color: '#718096', margin: 0 }}>Abhi tak koi order place nahi hua hai.</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ background: '#f2f2f2', textAlign: 'left' }}>
                          <th style={{ padding: '10px' }}>Order ID</th>
                          <th style={{ padding: '10px' }}>Customer</th>
                          <th style={{ padding: '10px' }}>Address</th>
                          <th style={{ padding: '10px' }}>Items</th>
                          <th style={{ padding: '10px' }}>Amount</th>
                          <th style={{ padding: '10px' }}>Payment</th>
                          <th style={{ padding: '10px' }}>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {orders.map((o) => (
                          <tr key={o.id} style={{ borderBottom: '1px solid #eee' }}>
                            <td style={{ padding: '10px', fontWeight: 'bold' }}>{o.id}</td>
                            <td style={{ padding: '10px' }}>{o.customerName}<br/><span style={{ fontSize: '11px', color: '#666' }}>{o.phone}</span></td>
                            <td style={{ padding: '10px', fontSize: '12px' }}>{o.address}</td>
                            <td style={{ padding: '10px' }}>{o.items.map(i => `${i.name} x${i.qty}`).join(', ')}</td>
                            <td style={{ padding: '10px', fontWeight: 'bold' }}>₹{o.totalAmount}</td>
                            <td style={{ padding: '10px' }}>{o.paymentMethod}</td>
                            <td style={{ padding: '10px' }}>
                              <span style={{ backgroundColor: '#c6f6d5', color: '#22543d', padding: '4px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>
                                {o.status}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}

            {/* 5. CUSTOMERS SECTION */}
            {adminTab === 'Customers' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Registered Customers ({customersList.length})</h2>
                
                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  {customersList.length === 0 ? (
                    <p style={{ color: '#718096', margin: 0 }}>Jab customer order karenge, tab unki profile yaha generate hogi.</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ background: '#f2f2f2', textAlign: 'left' }}>
                          <th style={{ padding: '10px' }}>Name</th>
                          <th style={{ padding: '10px' }}>Phone Number</th>
                          <th style={{ padding: '10px' }}>Total Orders</th>
                          <th style={{ padding: '10px' }}>Total Spent</th>
                        </tr>
                      </thead>
                      <tbody>
                        {customersList.map((c, i) => (
                          <tr key={i} style={{ borderBottom: '1px solid #eee' }}>
                            <td style={{ padding: '10px', fontWeight: 'bold' }}>{c.name}</td>
                            <td style={{ padding: '10px' }}>{c.phone}</td>
                            <td style={{ padding: '10px' }}>{c.totalOrders}</td>
                            <td style={{ padding: '10px', fontWeight: 'bold', color: '#0c831f' }}>₹{c.totalSpent}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}

            {/* 6. INVENTORY SECTION */}
            {adminTab === 'Inventory' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Inventory Stock Management</h2>
                
                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: '#f2f2f2', textAlign: 'left' }}>
                        <th style={{ padding: '10px' }}>Product Name</th>
                        <th style={{ padding: '10px' }}>Category</th>
                        <th style={{ padding: '10px' }}>Stock Available</th>
                        <th style={{ padding: '10px' }}>Stock Status</th>
                        <th style={{ padding: '10px' }}>Update Stock</th>
                      </tr>
                    </thead>
                    <tbody>
                      {products.map((p) => (
                        <tr key={p.id} style={{ borderBottom: '1px solid #eee' }}>
                          <td style={{ padding: '10px', fontWeight: 'bold' }}>{p.name}</td>
                          <td style={{ padding: '10px' }}>{p.category}</td>
                          <td style={{ padding: '10px', fontWeight: 'bold' }}>{p.stock}</td>
                          <td style={{ padding: '10px' }}>
                            {p.stock > 10 ? (
                              <span style={{ backgroundColor: '#c6f6d5', color: '#22543d', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>In Stock</span>
                            ) : p.stock > 0 ? (
                              <span style={{ backgroundColor: '#feebc8', color: '#742a2a', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>Low Stock</span>
                            ) : (
                              <span style={{ backgroundColor: '#fed7d7', color: '#9b2c2c', padding: '3px 8px', borderRadius: '12px', fontSize: '11px', fontWeight: 'bold' }}>Out of Stock</span>
                            )}
                          </td>
                          <td style={{ padding: '10px' }}>
                            <button onClick={() => setProducts(products.map(item => item.id === p.id ? {...item, stock: item.stock + 10} : item))} style={{ padding: '4px 8px', border: '1px solid #ccc', borderRadius: '4px', cursor: 'pointer', background: '#fff' }}>+10 Stock</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* 7. SALES REPORT SECTION */}
            {adminTab === 'Sales Report' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Sales Report Summary</h2>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '20px', marginBottom: '20px' }}>
                  <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                    <h4 style={{ margin: '0 0 5px', color: '#718096' }}>Total Revenue Generated</h4>
                    <h2 style={{ margin: 0, color: '#0c831f' }}>₹{totalSalesAmount}</h2>
                  </div>
                  <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                    <h4 style={{ margin: '0 0 5px', color: '#718096' }}>Completed Orders</h4>
                    <h2 style={{ margin: 0, color: '#3182ce' }}>{orders.length}</h2>
                  </div>
                </div>

                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <h3 style={{ marginTop: 0 }}>Sales Overview Log</h3>
                  {orders.length === 0 ? (
                    <p style={{ color: '#718096' }}>Koi sale generate hone par yaha sales history show hogi.</p>
                  ) : (
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ background: '#f2f2f2', textAlign: 'left' }}>
                          <th style={{ padding: '10px' }}>Date & Time</th>
                          <th style={{ padding: '10px' }}>Order ID</th>
                          <th style={{ padding: '10px' }}>Customer</th>
                          <th style={{ padding: '10px' }}>Amount</th>
                        </tr>
                      </thead>
                      <tbody>
                        {orders.map(o => (
                          <tr key={o.id} style={{ borderBottom: '1px solid #eee' }}>
                            <td style={{ padding: '10px' }}>{o.date}</td>
                            <td style={{ padding: '10px', fontWeight: 'bold' }}>{o.id}</td>
                            <td style={{ padding: '10px' }}>{o.customerName}</td>
                            <td style={{ padding: '10px', fontWeight: 'bold', color: '#0c831f' }}>₹{o.totalAmount}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}

            {/* 8. SETTINGS SECTION */}
            {adminTab === 'Settings' && (
              <div>
                <h2 style={{ margin: '0 0 20px', fontSize: '24px', color: '#1a202c' }}>Store Settings</h2>

                <div style={{ background: '#fff', padding: '20px', borderRadius: '8px', maxWidth: '500px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                  <div style={{ marginBottom: '15px' }}>
                    <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', marginBottom: '5px' }}>Store Name</label>
                    <input type="text" value={storeSettings.storeName} onChange={(e) => setStoreSettings({...storeSettings, storeName: e.target.value})} style={{ width: '100%', padding: '8px', border: '1px solid #ccc', borderRadius: '4px', boxSizing: 'border-box' }} />
                  </div>

                  <div style={{ marginBottom: '15px' }}>
                    <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', marginBottom: '5px' }}>Contact Phone</label>
                    <input type="text" value={storeSettings.phone} onChange={(e) => setStoreSettings({...storeSettings, phone: e.target.value})} style={{ width: '100%', padding: '8px', border: '1px solid #ccc', borderRadius: '4px', boxSizing: 'border-box' }} />
                  </div>

                  <div style={{ marginBottom: '15px' }}>
                    <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', marginBottom: '5px' }}>Store Location Address</label>
                    <input type="text" value={storeSettings.address} onChange={(e) => setStoreSettings({...storeSettings, address: e.target.value})} style={{ width: '100%', padding: '8px', border: '1px solid #ccc', borderRadius: '4px', boxSizing: 'border-box' }} />
                  </div>

                  <div style={{ marginBottom: '15px' }}>
                    <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', marginBottom: '5px' }}>Standard Delivery Fee (₹)</label>
                    <input type="number" value={storeSettings.deliveryFee} onChange={(e) => setStoreSettings({...storeSettings, deliveryFee: Number(e.target.value)})} style={{ width: '100%', padding: '8px', border: '1px solid #ccc', borderRadius: '4px', boxSizing: 'border-box' }} />
                  </div>

                  <div style={{ marginBottom: '20px' }}>
                    <label style={{ display: 'block', fontSize: '13px', fontWeight: 'bold', marginBottom: '5px' }}>Minimum Amount for Free Delivery (₹)</label>
                    <input type="number" value={storeSettings.minFreeDelivery} onChange={(e) => setStoreSettings({...storeSettings, minFreeDelivery: Number(e.target.value)})} style={{ width: '100%', padding: '8px', border: '1px solid #ccc', borderRadius: '4px', boxSizing: 'border-box' }} />
                  </div>

                  <button onClick={() => alert('Settings Saved Successfully!')} style={{ background: '#0c831f', color: '#fff', border: 'none', padding: '10px 20px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>Save Changes</button>
                </div>
              </div>
            )}

          </div>

        </div>
      )}

      {/* SLIDE-OVER CART & CHECKOUT DRAWER */}
      {isCartOpen && (
        <div style={{ position: 'fixed', top: 0, left: 0, width: '100vw', height: '100vh', backgroundColor: 'rgba(0, 0, 0, 0.5)', zIndex: 1000, display: 'flex', justifyContent: 'flex-end' }}>
          <div style={{ flex: 1 }} onClick={resetCheckoutFlow} />

          <div style={{ width: '100%', maxWidth: '420px', backgroundColor: '#f8f9fa', height: '100%', display: 'flex', flexDirection: 'column', boxShadow: '-4px 0 15px rgba(0,0,0,0.15)', position: 'relative' }}>
            
            {/* Header */}
            <div style={{ backgroundColor: '#fff', padding: '16px 20px', borderBottom: '1px solid #eee', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                {checkoutStep !== 'cart' && checkoutStep !== 'confirmed' && (
                  <button onClick={() => {
                    if (checkoutStep === 'address') setCheckoutStep('cart');
                    if (checkoutStep === 'payment') setCheckoutStep('address');
                    if (checkoutStep === 'summary') setCheckoutStep('payment');
                  }} style={{ background: 'none', border: 'none', fontSize: '18px', cursor: 'pointer' }}>←</button>
                )}
                <h3 style={{ margin: 0, fontSize: '18px', color: '#222' }}>
                  {checkoutStep === 'cart' && 'My Cart'}
                  {checkoutStep === 'address' && 'Delivery Address'}
                  {checkoutStep === 'payment' && 'Payment Method'}
                  {checkoutStep === 'summary' && 'Order Summary'}
                  {checkoutStep === 'confirmed' && 'Order Confirmed!'}
                </h3>
              </div>
              <button onClick={resetCheckoutFlow} style={{ background: 'none', border: 'none', fontSize: '20px', cursor: 'pointer', color: '#888' }}>✕</button>
            </div>

            {cart.length > 0 && checkoutStep !== 'confirmed' && (
              <div style={{ backgroundColor: '#eef9f1', color: '#0c831f', padding: '10px 20px', fontSize: '13px', fontWeight: 'bold' }}>
                ⚡ Delivery in 10 minutes to {storeSettings.address}
              </div>
            )}

            {/* Scrollable Area */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '16px 20px' }}>

              {/* EMPTY CART */}
              {cart.length === 0 && checkoutStep !== 'confirmed' && (
                <div style={{ textAlign: 'center', padding: '60px 20px' }}>
                  <div style={{ fontSize: '50px', marginBottom: '10px' }}>🛒</div>
                  <h4 style={{ margin: '0 0 8px', color: '#333' }}>Your Cart is Empty</h4>
                  <p style={{ fontSize: '13px', color: '#777', marginBottom: '20px' }}>Add items to start shopping!</p>
                  <button onClick={resetCheckoutFlow} style={{ backgroundColor: '#0c831f', color: '#fff', border: 'none', padding: '10px 20px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer' }}>Start Shopping</button>
                </div>
              )}

              {/* STEP 1: CART & BILL */}
              {cart.length > 0 && checkoutStep === 'cart' && (
                <div>
                  <div style={{ background: '#fff', borderRadius: '10px', padding: '10px 15px', marginBottom: '15px' }}>
                    {cart.map(item => (
                      <div key={item.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 0', borderBottom: '1px solid #f0f0f0' }}>
                        <div>
                          <div style={{ fontWeight: '600', fontSize: '14px', color: '#222' }}>{item.name}</div>
                          <div style={{ fontSize: '12px', color: '#888' }}>{item.unit} | ₹{item.price}</div>
                        </div>

                        {/* Quantity Controls */}
                        <div style={{ display: 'flex', alignItems: 'center', border: '1px solid #0c831f', borderRadius: '6px', overflow: 'hidden' }}>
                          <button onClick={() => updateQuantity(item.id, -1)} style={{ backgroundColor: '#eef9f1', color: '#0c831f', border: 'none', padding: '4px 10px', cursor: 'pointer', fontWeight: 'bold' }}>-</button>
                          <span style={{ padding: '0 10px', fontSize: '13px', fontWeight: 'bold' }}>{item.qty}</span>
                          <button onClick={() => updateQuantity(item.id, 1)} style={{ backgroundColor: '#eef9f1', color: '#0c831f', border: 'none', padding: '4px 10px', cursor: 'pointer', fontWeight: 'bold' }}>+</button>
                        </div>
                      </div>
                    ))}
                  </div>

                  {/* Bill Summary */}
                  <div style={{ background: '#fff', borderRadius: '10px', padding: '15px' }}>
                    <h4 style={{ margin: '0 0 12px', fontSize: '15px', color: '#333' }}>Bill Summary</h4>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: '#555', marginBottom: '8px' }}>
                      <span>Item Total</span><span>₹{itemsTotal}</span>
                    </div>
                    {discount > 0 && (
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: '#0c831f', marginBottom: '8px' }}>
                        <span>Discount</span><span>-₹{discount}</span>
                      </div>
                    )}
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: '#555', marginBottom: '8px' }}>
                      <span>Delivery Charge</span><span>{deliveryFee === 0 ? <strong style={{ color: '#0c831f' }}>FREE</strong> : `₹${deliveryFee}`}</span>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', color: '#555', marginBottom: '12px' }}>
                      <span>Platform Fee</span><span>₹{platformFee}</span>
                    </div>
                    <hr style={{ border: 'none', borderTop: '1px dashed #ddd', margin: '10px 0' }} />
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '15px', fontWeight: 'bold', color: '#111' }}>
                      <span>To Pay</span><span>₹{grandTotal}</span>
                    </div>
                  </div>
                </div>
              )}

              {/* STEP 2: ADDRESS */}
              {checkoutStep === 'address' && (
                <div style={{ background: '#fff', borderRadius: '10px', padding: '15px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                    <h4 style={{ margin: 0, fontSize: '15px' }}>Delivery Address</h4>
                    <button onClick={() => setIsEditingAddress(!isEditingAddress)} style={{ color: '#0c831f', background: 'none', border: 'none', fontWeight: 'bold', cursor: 'pointer', fontSize: '12px' }}>
                      {isEditingAddress ? 'Save' : 'Change'}
                    </button>
                  </div>

                  {!isEditingAddress ? (
                    <div style={{ fontSize: '13px', color: '#444', lineHeight: '1.5' }}>
                      <p style={{ margin: 0, fontWeight: 'bold' }}>{address.fullName} ({address.phone})</p>
                      <p style={{ margin: '4px 0 0' }}>{address.houseNo}, {address.area}</p>
                      <p style={{ margin: '2px 0 0', color: '#777' }}>{address.city}, Maharashtra - {address.pincode}</p>
                    </div>
                  ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      <input type="text" value={address.fullName} onChange={e => setAddress({...address, fullName: e.target.value})} placeholder="Full Name" style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                      <input type="text" value={address.phone} onChange={e => setAddress({...address, phone: e.target.value})} placeholder="Phone Number" style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                      <input type="text" value={address.houseNo} onChange={e => setAddress({...address, houseNo: e.target.value})} placeholder="House / Shop No." style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                      <input type="text" value={address.area} onChange={e => setAddress({...address, area: e.target.value})} placeholder="Area / Street" style={{ padding: '8px', border: '1px solid #ccc', borderRadius: '4px' }} />
                    </div>
                  )}
                </div>
              )}

              {/* STEP 3: PAYMENT */}
              {checkoutStep === 'payment' && (
                <div style={{ background: '#fff', borderRadius: '10px', padding: '15px' }}>
                  <h4 style={{ margin: '0 0 15px', fontSize: '15px' }}>Payment Method</h4>
                  {[
                    { id: 'upi', name: 'UPI (GPay / PhonePe / Paytm)', icon: '📱' },
                    { id: 'cod', name: 'Cash on Delivery (COD)', icon: '💵' },
                    { id: 'card', name: 'Debit / Credit Card', icon: '💳' }
                  ].map(method => (
                    <div 
                      key={method.id}
                      onClick={() => setPaymentMethod(method.id)}
                      style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px', borderRadius: '8px', border: paymentMethod === method.id ? '2px solid #0c831f' : '1px solid #e0e0e0', marginBottom: '10px', cursor: 'pointer', backgroundColor: paymentMethod === method.id ? '#eef9f1' : '#fff' }}
                    >
                      <span style={{ fontSize: '14px', fontWeight: '500' }}>{method.icon} {method.name}</span>
                      <input type="radio" checked={paymentMethod === method.id} readOnly />
                    </div>
                  ))}
                </div>
              )}

              {/* STEP 4: SUMMARY */}
              {checkoutStep === 'summary' && (
                <div>
                  <div style={{ background: '#fff', borderRadius: '10px', padding: '12px 15px', marginBottom: '12px' }}>
                    <h5 style={{ margin: '0 0 6px', color: '#888', fontSize: '12px' }}>DELIVERING TO</h5>
                    <div style={{ fontSize: '13px', fontWeight: 'bold' }}>{address.fullName}</div>
                    <div style={{ fontSize: '12px', color: '#555' }}>{address.houseNo}, {address.area}, Kinwat</div>
                  </div>

                  <div style={{ background: '#fff', borderRadius: '10px', padding: '12px 15px', marginBottom: '12px' }}>
                    <h5 style={{ margin: '0 0 6px', color: '#888', fontSize: '12px' }}>PAYMENT METHOD</h5>
                    <div style={{ fontSize: '13px', fontWeight: 'bold', textTransform: 'uppercase' }}>{paymentMethod}</div>
                  </div>

                  <div style={{ background: '#fff', borderRadius: '10px', padding: '12px 15px' }}>
                    <h5 style={{ margin: '0 0 8px', color: '#888', fontSize: '12px' }}>ITEMS ({cart.length})</h5>
                    {cart.map(item => (
                      <div key={item.id} style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', marginBottom: '4px' }}>
                        <span>{item.name} x {item.qty}</span>
                        <span>₹{item.price * item.qty}</span>
                      </div>
                    ))}
                    <hr style={{ border: 'none', borderTop: '1px solid #eee', margin: '8px 0' }} />
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 'bold', fontSize: '14px' }}>
                      <span>Total Amount</span>
                      <span>₹{grandTotal}</span>
                    </div>
                  </div>
                </div>
              )}

              {/* STEP 5: CONFIRMED */}
              {checkoutStep === 'confirmed' && (
                <div style={{ textAlign: 'center', padding: '40px 10px' }}>
                  <div style={{ fontSize: '60px', marginBottom: '10px' }}>🎉</div>
                  <h3 style={{ color: '#0c831f', margin: '0 0 8px' }}>Order Placed Successfully!</h3>
                  <p style={{ fontSize: '13px', color: '#666', marginBottom: '20px' }}>Your order will arrive in 10 minutes.</p>
                  <button onClick={() => { setCart([]); resetCheckoutFlow(); }} style={{ backgroundColor: '#0c831f', color: '#fff', border: 'none', padding: '12px 24px', borderRadius: '6px', fontWeight: 'bold', cursor: 'pointer', width: '100%' }}>Back to Store</button>
                </div>
              )}

            </div>

            {/* Bottom Button */}
            {cart.length > 0 && checkoutStep !== 'confirmed' && (
              <div style={{ backgroundColor: '#fff', padding: '12px 20px', borderTop: '1px solid #eee' }}>
                <button
                  onClick={() => {
                    if (checkoutStep === 'cart') setCheckoutStep('address');
                    else if (checkoutStep === 'address') setCheckoutStep('payment');
                    else if (checkoutStep === 'payment') setCheckoutStep('summary');
                    else if (checkoutStep === 'summary') handlePlaceOrder();
                  }}
                  style={{ backgroundColor: '#0c831f', color: '#fff', border: 'none', width: '100%', padding: '14px', borderRadius: '8px', fontSize: '15px', fontWeight: 'bold', display: 'flex', justifyContent: 'space-between', cursor: 'pointer' }}
                >
                  <span>₹{grandTotal}</span>
                  <span>
                    {checkoutStep === 'cart' && 'Proceed to Address →'}
                    {checkoutStep === 'address' && 'Select Payment →'}
                    {checkoutStep === 'payment' && 'Review Order →'}
                    {checkoutStep === 'summary' && 'Place Order →'}
                  </span>
                </button>
              </div>
            )}

          </div>
        </div>
      )}

    </div>
  );
}

export default App;