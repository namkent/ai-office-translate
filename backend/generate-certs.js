const forge = require('node-forge');
const fs = require('fs');
const path = require('path');

function generateCerts() {
  const certPath = path.join(__dirname, 'cert.pem');
  const keyPath = path.join(__dirname, 'key.pem');

  // Kiểm tra xem chứng chỉ đã tồn tại chưa
  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    console.log('SSL certificates already exist. Skipping generation.');
    return;
  }

  console.log('Generating self-signed SSL certificates for localhost...');

  // Tạo cặp khóa RSA 2048-bit
  const keys = forge.pki.rsa.generateKeyPair(2048);

  // Tạo chứng chỉ
  const cert = forge.pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = '01' + Date.now();
  cert.validity.notBefore = new Date();
  cert.validity.notAfter = new Date();
  cert.validity.notAfter.setFullYear(cert.validity.notBefore.getFullYear() + 1); // Hiệu lực 1 năm

  const attrs = [
    { name: 'commonName', value: 'localhost' },
    { name: 'countryName', value: 'VN' },
    { name: 'stateOrProvinceName', value: 'Hanoi' },
    { name: 'localityName', value: 'Hanoi' },
    { name: 'organizationName', value: 'AI Office Translate' },
    { name: 'organizationalUnitName', value: 'Development' }
  ];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);

  // Thiết lập các Extension (Bắt buộc đối với chứng chỉ localhost ở các trình duyệt hiện đại)
  cert.setExtensions([
    {
      name: 'basicConstraints',
      cA: true
    },
    {
      name: 'keyUsage',
      keyCertSign: true,
      digitalSignature: true,
      nonRepudiation: true,
      keyEncipherment: true,
      dataEncipherment: true
    },
    {
      name: 'extKeyUsage',
      serverAuth: true,
      clientAuth: true
    },
    {
      name: 'subjectAltName',
      altNames: [
        { type: 2, value: 'localhost' }, // DNS:localhost
        { type: 7, ip: '127.0.0.1' }     // IP:127.0.0.1
      ]
    }
  ]);

  // Ký chứng chỉ bằng private key
  cert.sign(keys.privateKey, forge.md.sha256.create());

  // Chuyển đổi sang định dạng PEM
  const pemCert = forge.pki.certificateToPem(cert);
  const pemKey = forge.pki.privateKeyToPem(keys.privateKey);

  // Đảm bảo thư mục tồn tại
  const dir = path.dirname(certPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  // Ghi tệp
  fs.writeFileSync(keyPath, pemKey);
  fs.writeFileSync(certPath, pemCert);

  console.log('SSL Certificates generated successfully!');
  console.log(`- Private Key: ${keyPath}`);
  console.log(`- Certificate: ${certPath}`);
}

// Chạy trực tiếp nếu script được gọi từ dòng lệnh
if (require.main === module) {
  generateCerts();
}

module.exports = generateCerts;
