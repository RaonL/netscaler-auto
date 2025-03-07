const NodeSSH = require('node-ssh');
const CronJob = require('cron').CronJob;
const fs = require('fs');
const ExcelJS = require('exceljs');

const ssh = new NodeSSH();

// NetScaler 연결 설정
const config = {
  host: 'your_netscaler_ip',
  username: 'your_username',
  password: 'your_password'
};

// 로그 파일 설정
const logFile = 'netscaler_check.log';
const excelFile = 'netscaler_status.xlsx';

// 로그 함수
function log(message) {
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] ${message}\n`;
  fs.appendFile(logFile, logMessage, (err) => {
    if (err) {
      console.error('Failed to write to log file:', err);
    }
  });
  console.log(logMessage);
}

// Excel 파일 생성 및 설정
async function createExcelFile() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('NetScaler Status');

  // 헤더 추가
  sheet.addRow(['Timestamp', 'CPU Usage', 'Memory Usage', 'Fan Status', 'Power Status', 'Receive Bytes', 'Transmit Bytes', 'HDD Status', 'LB Status']);

  await workbook.xlsx.writeFile(excelFile);
  log('Excel file created.');
}

// Excel 파일에 데이터 추가
async function addDataToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelFile);
  const sheet = workbook.getWorksheet('NetScaler Status');

  sheet.addRow(data);

  await workbook.xlsx.writeFile(excelFile);
  log('Data added to Excel file.');
}

// NetScaler에 SSH 연결
async function connectToNetScaler() {
  try {
    await ssh.connect(config);
    log('Connected to NetScaler');
    return true;
  } catch (err) {
    log(`Failed to connect to NetScaler: ${err}`);
    return false;
  }
}

// CPU 사용률 확인
async function checkCPU() {
  try {
    const result = await ssh.execCommand('show cpu');
    if (result.stderr) {
      log(`CPU check error: ${result.stderr}`);
      return null;
    }
    const cpuUsage = result.stdout.match(/CPU usage:\s*(\d+\.\d+)%/)?.[1];
    log(`CPU Usage: ${cpuUsage}%`);
    return cpuUsage;
  } catch (err) {
    log(`CPU check error: ${err}`);
    return null;
  }
}

// 메모리 사용률 확인
async function checkMemory() {
  try {
    const result = await ssh.execCommand('show memory');
    if (result.stderr) {
      log(`Memory check error: ${result.stderr}`);
      return null;
    }
    const memoryUsage = result.stdout.match(/Memory usage:\s*(\d+)%/)?.[1];
    log(`Memory Usage: ${memoryUsage}%`);
    return memoryUsage;
  } catch (err) {
    log(`Memory check error: ${err}`);
    return null;
  }
}

// 팬 상태 확인 (이 부분은 NetScaler 모델에 따라 명령어가 다를 수 있습니다.)
async function checkFanStatus() {
  try {
    const result = await ssh.execCommand('show hardware');
    if (result.stderr) {
      log(`Fan check error: ${result.stderr}`);
      return 'Error';
    }
    // show hardware 명령어 결과에서 팬 상태를 파싱하는 로직 추가
    const fanStatus = result.stdout.includes('Fan Speed') ? 'OK' : 'Error';
    log(`Fan Status: ${fanStatus}`);
    return fanStatus;
  } catch (err) {
    log(`Fan check error: ${err}`);
    return 'Error';
  }
}

// 전원 상태 확인 (이 부분은 NetScaler 모델에 따라 명령어가 다를 수 있습니다.)
async function checkPowerStatus() {
  try {
    const result = await ssh.execCommand('show hardware');
    if (result.stderr) {
      log(`Power check error: ${result.stderr}`);
      return 'Error';
    }
    // show hardware 명령어 결과에서 전원 상태를 파싱하는 로직 추가
    const powerStatus = result.stdout.includes('Power Supply') ? 'OK' : 'Error';
    log(`Power Status: ${powerStatus}`);
    return powerStatus;
  } catch (err) {
    log(`Power check error: ${err}`);
    return 'Error';
  }
}

// 처리량 확인
async function checkThroughput() {
  try {
    const result = await ssh.execCommand('show interface stats');
    if (result.stderr) {
      log(`Throughput check error: ${result.stderr}`);
      return { rxBytes: null, txBytes: null };
    }

    // 정규 표현식을 사용하여 처리량 관련 정보 추출
    const throughputMatch = result.stdout.match(/RxBytes:\s*(\d+)\s*TxBytes:\s*(\d+)/);

    if (throughputMatch && throughputMatch.length === 3) {
      const rxBytes = parseInt(throughputMatch[1]);
      const txBytes = parseInt(throughputMatch[2]);
      log(`Receive Bytes: ${rxBytes} bytes, Transmit Bytes: ${txBytes} bytes`);
      return { rxBytes, txBytes };
    } else {
      log('Throughput information not found.');
      return { rxBytes: null, txBytes: null };
    }
  } catch (err) {
    log(`Throughput check error: ${err}`);
    return { rxBytes: null, txBytes: null };
  }
}

// HDD 상태 확인 (이 부분은 NetScaler 모델에 따라 명령어가 다를 수 있습니다.)
async function checkHDDStatus() {
  try {
    const result = await ssh.execCommand('df -h');
    if (result.stderr) {
      log(`HDD check error: ${result.stderr}`);
      return 'Error';
    }
    // df -h 명령어 결과에서 HDD 상태를 파싱하는 로직 추가
    const hddStatus = result.stdout.includes('/dev/sda1') ? 'OK' : 'Error';
    log(`HDD Status: ${hddStatus}`);
    return hddStatus;
  } catch (err) {
    log(`HDD check error: ${err}`);
    return 'Error';
  }
}

// LB 상태 확인 (이 부분은 실제 LB 상태를 확인하는 명령어로 대체해야 합니다.)
async function checkLBStatus() {
  try {
    const result = await ssh.execCommand('show serviceGroup');
    if (result.stderr) {
      log(`LB check error: ${result.stderr}`);
      return 'Error';
    }
    // show serviceGroup 명령어 결과에서 LB 상태를 파싱하는 로직 추가
    const lbStatus = result.stdout.includes('STATE : ENABLED') ? 'OK' : 'Error';
    log(`LB Status: ${lbStatus}`);
    return lbStatus;
  } catch (err) {
    log(`LB check error: ${err}`);
    return 'Error';
  }
}

// 모든 점검 항목 실행
async function runChecks() {
  if (await connectToNetScaler()) {
    const timestamp = new Date().toISOString();
    const cpuUsage = await checkCPU();
    const memoryUsage = await checkMemory();
    const fanStatus = await checkFanStatus();
    const powerStatus = await checkPowerStatus();
    const throughput = await checkThroughput();
    const hddStatus = await checkHDDStatus();
    const lbStatus = await checkLBStatus();

    const data = [timestamp, cpuUsage, memoryUsage, fanStatus, powerStatus, throughput.rxBytes, throughput.txBytes, hddStatus, lbStatus];
    await addDataToExcel(data);

    ssh.dispose();
  }
}

// Cron 스케줄러 설정 (매 5분마다 실행)
async function startMonitoring() {
  await createExcelFile(); // Excel 파일 생성
  const job = new CronJob('*/5 * * * *', function() {
    log('Running NetScaler checks...');
    runChecks();
  }, null, true, 'America/Los_Angeles');
  job.start();

  log('NetScaler auto check program started.');
}

startMonitoring();
