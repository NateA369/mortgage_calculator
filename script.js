document.addEventListener('DOMContentLoaded', function() {
    // Function to embed Excel spreadsheet
    function embedExcelSpreadsheet() {
        const excelContainer = document.getElementById('excel-container');
        excelContainer.innerHTML = '<iframe src="https://1drv.ms/x/c/746d7f6f4436bc28/IQQihmhHM3qMSrA3nhJRvwp5ARXnOoQyK-py3EmPDP7ZRHc?embed=true&wdAllowInteractivity=False&AllowTyping=True&ActiveCell=\'Sheet1\'!D5&wdDownloadButton=True&wdInConfigurator=True" width="100%" height="600" frameborder="0" scrolling="no"></iframe>';
    }
    // Call the function to embed the spreadsheet
    embedExcelSpreadsheet();
});