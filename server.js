const express = require("express");
const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = 3000;
const app = express();

app.use(express.static(__dirname));

// CORS for Office add-ins
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "*");
  next();
});

// Try to use HTTPS with self-signed cert
const certPath = path.join(__dirname, "certs", "localhost.crt");
const keyPath = path.join(__dirname, "certs", "localhost.key");

if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
  https
    .createServer(
      {
        cert: fs.readFileSync(certPath),
        key: fs.readFileSync(keyPath),
      },
      app
    )
    .listen(PORT, () => {
      console.log(`\n  ✅ HTTPS 服务器已启动`);
      console.log(`  🔗 https://localhost:${PORT}`);
      console.log(`  📋 在 Word 中加载: https://localhost:${PORT}/manifest.xml\n`);
    });
} else {
  // Generate self-signed cert using PowerShell
  console.log("正在生成自签名证书...");
  const certDir = path.join(__dirname, "certs");
  if (!fs.existsSync(certDir)) fs.mkdirSync(certDir);

  try {
    const { execSync } = require("child_process");
    // Generate cert and key using openssl-style via PowerShell
    const psCmd = [
      `$cert = New-SelfSignedCertificate`,
      `-DnsName "localhost"`,
      `-CertStoreLocation "cert:\\LocalMachine\\My"`,
      `-NotAfter (Get-Date).AddYears(5)`,
      `-KeyExportPolicy Exportable`,
      `-Type SSLServerAuthentication`,
      `; $pwd = ConvertTo-SecureString -String "temp" -Force -AsPlainText`,
      `; Export-PfxCertificate -Cert $cert -FilePath "${certDir}\\localhost.pfx" -Password $pwd | Out-Null`,
    ].join(" ");

    execSync(`powershell -Command "${psCmd}"`, { stdio: "pipe" });
    console.log("PFX 证书已生成，正在转换为 PEM...");

    // Try openssl to convert
    try {
      execSync(
        `openssl pkcs12 -in "${certDir}\\localhost.pfx" -out "${certDir}\\localhost.crt" -nokeys -passin pass:temp -passout pass:temp`,
        { stdio: "pipe" }
      );
      execSync(
        `openssl pkcs12 -in "${certDir}\\localhost.pfx" -out "${certDir}\\localhost.key" -nocerts -nodes -passin pass:temp`,
        { stdio: "pipe" }
      );
    } catch (e) {
      console.log("\n⚠️  需要手动转换证书，请安装 openssl 或手动操作:");
      console.log(`   openssl pkcs12 -in ${certDir}\\localhost.pfx -out ${certDir}\\localhost.crt -nokeys -passin pass:temp`);
      console.log(`   openssl pkcs12 -in ${certDir}\\localhost.pfx -out ${certDir}\\localhost.key -nocerts -nodes -passin pass:temp\n`);
    }
  } catch (e) {
    console.log("证书生成失败，使用 HTTP 模式");
  }

  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    https
      .createServer(
        {
          cert: fs.readFileSync(certPath),
          key: fs.readFileSync(keyPath),
        },
        app
      )
      .listen(PORT, () => {
        console.log(`\n  ✅ HTTPS 服务器已启动`);
        console.log(`  🔗 https://localhost:${PORT}`);
        console.log(`  📋 在 Word 中加载: https://localhost:${PORT}/manifest.xml\n`);
      });
  } else {
    app.listen(PORT, () => {
      console.log(`\n  ⚠️  HTTP 服务器已启动 (Word插件可能需要HTTPS)`);
      console.log(`  🔗 http://localhost:${PORT}\n`);
    });
  }
}
