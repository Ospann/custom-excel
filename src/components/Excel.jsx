import generateStyledExcel from "../helper/Excel";

const StyledExcelDownload = ({ data }) => {
    const handleDownload = async () => {
        const excelBuffer = await generateStyledExcel(data);
        const blob = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const url = URL.createObjectURL(blob);
        const downloadLink = document.createElement("a");
        downloadLink.href = url;
        downloadLink.download = "styled-excel-file.xlsx"; // Change the filename as needed
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    };

    return (
        <button onClick={handleDownload}>Download Styled Excel</button>
    );
};

export default StyledExcelDownload;
