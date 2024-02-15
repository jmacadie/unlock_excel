// The path to the vba file within an xlsx or xlsb file
pub const ZIP_VBA_PATH: &str = "xl/vbaProject.bin";

// The path to the vba project stream within an xls file
pub const CFB_VBA_PATH: &str = "/_VBA_PROJECT_CUR/PROJECT";

// The path to the project stream within a VBA compound file
pub const PROJECT_PATH: &str = "/PROJECT";

// The project properties of an unlocked project
pub const UNLOCKED_ID: &str = "ID=\"{3C6F1B8B-BDBE-4F1B-AA02-BCA23D695691}\"\r\n";
pub const UNLOCKED_CMG: &str = "CMG=\"1E1C02263E5A585E585E585E585E\"\r\n";
pub const UNLOCKED_DPB: &str = "DPB=\"3C3E2044206321632163\"\r\n";
pub const UNLOCKED_GC: &str = "GC=\"5A58466A656B656B9A\"\r\n";
