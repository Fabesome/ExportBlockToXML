using System;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.AddIn;
using Siemens.Engineering.AddIn.Menu;
using System.Linq;
using System.Windows.Forms;



namespace AddInExample
{
    public class AddIn : ContextMenuAddIn
    {
        private readonly TiaPortal _tiaPortal;
        private readonly Settings _settings;

        public AddIn(TiaPortal tiaPortal) : base("Export Block To XML")
        {
            _tiaPortal = tiaPortal;
            _settings = Settings.Load();
        }

        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            //Export Menus
            addInRootSubmenu.Items.AddActionItem<FB>("Export Block to XML", OnClick);
            addInRootSubmenu.Items.AddActionItem<FC>("Export Block to XML", OnClick);
            addInRootSubmenu.Items.AddActionItem<DataBlock>("Export Block to XML", OnClick);
            addInRootSubmenu.Items.AddActionItem<PlcTagTable>("Export List to XML", OnClickTagList);

            //Check Menu for DBs
            //addInRootSubmenu.Items.AddActionItem<DataBlock>("Check fault-texts in DB", OnClickCheck);



            //addInRootSubmenu.Items.AddActionItem<FB>("Convert to FC", AddInClick);
            addInRootSubmenu.Items.AddActionItem<IEngineeringObject>("Please select only FBs/FCs/DBs/TagLists", menuSelectionProvider => { }, InfoTextStatus);

            //Submenu settingsSubmenu = addInRootSubmenu.Items.AddSubmenu("Settings");
            //settingsSubmenu.Items.AddActionItemWithCheckBox<IEngineeringObject>("Check Box", _settings.CheckBoxOnClick, _settings.CheckBoxDisplayStatus);
            //settingsSubmenu.Items.AddActionItemWithRadioButton<IEngineeringObject>("Radio Button 1", _settings.RadioButton1OnClick, _settings.RadioButton1DisplayStatus);
            //settingsSubmenu.Items.AddActionItemWithRadioButton<IEngineeringObject>("Radio Button 2", _settings.RadioButton2OnClick, _settings.RadioButton2DisplayStatus);
        }

        private void OnClick(MenuSelectionProvider menuSelectionProvider)
        {
            // TODO: Replace this with your own on click logic

            string projectName = _tiaPortal.Projects.First(project => project.IsPrimary).Name;
            List<IEngineeringObject> selectedObjects = menuSelectionProvider.GetSelection<IEngineeringObject>().ToList();

            //FolderBrowserDialog fbd = new FolderBrowserDialog();
            //fbd.Description = "Select save path";
            //fbd.ShowDialog();

            foreach (PlcBlock x in selectedObjects) {

                    FileInfo f = new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + x.Name + ".xml");
                    //  FileStream fs = new FileStream(fbd.SelectedPath, FileMode.Create);
                    x.Export(f, ExportOptions.WithDefaults | ExportOptions.WithReadOnly);
            }

        }
        private void OnClickTagList(MenuSelectionProvider menuSelectionProvider)
        {
            // TODO: Replace this with your own on click logic

            string projectName = _tiaPortal.Projects.First(project => project.IsPrimary).Name;
            List<IEngineeringObject> selectedObjects = menuSelectionProvider.GetSelection<IEngineeringObject>().ToList();

            //FolderBrowserDialog fbd = new FolderBrowserDialog();
            //fbd.Description = "Select save path";
            //fbd.ShowDialog();

            foreach (PlcTagTable x in selectedObjects)
            {

                FileInfo f = new FileInfo(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + x.Name + ".xml");
                //  FileStream fs = new FileStream(fbd.SelectedPath, FileMode.Create);
                x.Export(f, ExportOptions.None/*, ExportOptions.WithDefaults | ExportOptions.WithReadOnly*/);
            }

        }

        private void OnClickCheck(MenuSelectionProvider menuSelectionProvider)
        {
            string projectName = _tiaPortal.Projects.First(project => project.IsPrimary).Name;
            List<IEngineeringObject> selectedObjects = menuSelectionProvider.GetSelection<IEngineeringObject>().ToList();

            //init languages
            
            //english
            Language englishLanguage = _tiaPortal.Projects.First(project => project.IsPrimary).LanguageSettings.Languages.Find(new CultureInfo("enUS"));
            //german 
            Language germanLanguage = _tiaPortal.Projects.First(project => project.IsPrimary).LanguageSettings.Languages.Find(new CultureInfo("deDE"));

            //MultilingualText comment = _tiaPortal.Projects.First(project => project.IsPrimary).Comment;
            //MultilingualTextItemComposition mltItemComposition = comment.Items;
            //MultilingualTextItem englishComment = mltItemComposition.Find(englishLanguage);
            //englishComment.Text = "English comment";

        }

            private static MenuStatus InfoTextStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            var show = false;

            foreach (IEngineeringObject engineeringObject in menuSelectionProvider.GetSelection())
            {
                if (!(engineeringObject.GetType() == menuSelectionProvider.GetSelection().First().GetType() && ((engineeringObject is DataBlock) || (engineeringObject is FB) || (engineeringObject is FC) || (engineeringObject is PlcTagTable))))
                {
                    show = true;
                    break;
                }
            }

            return show ? MenuStatus.Disabled : MenuStatus.Hidden;
        }

        private MenuStatus DisplayStatus(MenuSelectionProvider<IEngineeringObject> menuSelectionProvider)
        {
            return MenuStatus.Enabled;
        }
    }
}
