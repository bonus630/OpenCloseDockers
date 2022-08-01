using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using corel = Corel.Interop.VGCore;
using System.Collections.Specialized;
using System.Diagnostics;

namespace OpenCloseDockers
{

    public partial class ControlUI : UserControl
    {
        private corel.Application corelApp;
        private Styles.StylesController stylesController;


        //Guid dockers list for CorelDraw 2020
        private readonly string[] dockersGuids = new string[]{
            "328c33c0-b5b9-b7b5-4952-df2535639ab0",
            "0938ed71-aa9c-9fbc-440f-524d043745d3",
            "1c7ef669-9f29-c1b4-4985-045b70f9e58c",
            "0750d8df-2c52-43d7-9aa0-9ec2531dfe42",
            "a8b139f6-dd44-0883-49e9-8b18d2b16f38",
            "4a8feb78-844a-429e-8797-387c002a359d",
            "28b4a595-af12-4199-aa2d-2d8e9a008b60",
            "b3fb1964-9f5d-4650-8e4e-20b88bb9671c",
            "1a0af10a-c0c8-fead-4f94-fb89f4fc2db2",
            "e6873daf-b5b9-4957-bb78-3335fc8c6054",
            "de52917e-e529-456c-a9c0-dbb13838ef85",
            "83f56d75-2bdd-4fb7-91db-b38894f5684d",
            "454fda4e-4634-4079-8534-17797a64dfae",
            "18d6991b-8ef8-4b48-be1a-d55f769648a9",
            "7ef910da-e881-4f59-86ea-94456da7761d",
            "02bdc6f4-065d-9fa6-44c2-a41ddf4d3b10",
            "799476ff-2b80-43cb-b51a-a31ac9141618",
            "73675757-b6d9-44fb-87b8-f86089d94e87",
            "30afc88c-1e67-4b21-9893-73cbb3895bd3",
            "a2f385ec-72c6-4364-bead-4cd09e6dda93",
            "e79e3cf2-c93d-4694-b463-f2f7113872e8",
            "899d052c-fc8b-93b0-43c7-bda85ca8f417",
            "d38b764b-2616-7bab-437b-2233d902aec5",
            "bf87b90f-5a78-40b6-8451-d65ae17ed14b",
            "370e97bf-651f-4431-9bfd-140035be0c7e",
            "155feb2c-56dc-4a8c-9362-9ab38561b267",
            "ce35466c-2377-4d42-b293-e76e290e2199",
            "b9fa0ab0-e60e-38a0-443d-37f0488d6aa8",
            "2191be05-845e-4ee1-8d61-c036f492fee8",
            "9452d7fb-1204-4e0f-9ff6-92b7806c199d",
            "7b49396f-f989-c087-4ed9-de2b835dcf68",
            "dbd626b4-ed44-3ca1-4ce8-71921595bd74",
            "9d5d7bc9-3227-1195-4c7a-fe91788e2367",
            "102052d9-0a87-481a-af3b-00f1bce81f9f",
            "8151bb6c-2ec0-84b7-4d1e-a379eabd0156",
            "2336feaf-4a7c-6988-4338-faf5a2c348da",
            "13ab367b-038e-aebb-4f5e-0540a26cdc59",
            "f3ab7212-445a-c7b4-47db-94a7d944ccda",
            "d632ed56-5e70-bca1-43f8-249a6342dbf2",
            "aed09712-12ac-d98764723-aa3167ef73ca",
            "5e850daf-af40-d480-4201-10e9cf35b841",
            "dbda185b-ac02-f6b4-4052-654348a46a39",
            "a3f597ca-e079-ab88-4e2f-28a3b4db6d8e",
            "11b86ad8-6f3a-451c-9874-5799cd293e6b",
            "4432782c-684b-497d-b35f-de0b0d9a8e19",
            "7a17fa29-04e0-4dc1-ab7c-7062f026b7f7",
            "af155c35-0e67-4ee8-be99-fdb2b356d92d",
            "6af91e05-fe18-47b4-a835-256278b92439",
            "a314932f-cdd1-435c-a57c-210395d49e42",
            "de4676f4-cb4a-4ba2-b008-8205e3ec5f84",
            "c50a493c-e244-48b3-9d15-accf12ebc4c4",
            "15d11cd3-fb72-47a2-a7f9-e82c6be07596",
            "67b8ac81-501a-4ffc-a154-9455a3dce765",
            "77a1cd97-3947-4b88-87ad-405a03aa4f01",
            "9a4ffb50-8e3b-4ab0-ab7b-0ad322a5b459",
            "2b40530d-af10-4ea1-a607-837ddebec17b",
            "76baddcc-6deb-e481-42c6-389833de8d5a",
            "957f5aab-d8f9-35aa-4f2c-20c938c7c08b",
            "cd832eb6-02cf-6999-4500-13a0e1de5a9b",
            "3285915b-5668-9986-4e1e-a4612906ca11",
            "4ee2c250-2fb6-1cb3-4b55-f03d2832988b",
            "d435185d-b276-09a9-4fb2-0a26d97a27f9",
            "76853e70-686d-838e-4ddc-b305bf5369a6",
            "ed27b414-58cd-9b82-447b-61f9d4e27819",
            "1689e79a-e8f6-4894-4b74-051309aebcbf",
            "6f720125-b417-6095-4e53-3356118dd82b",
            "41e2e0cc-0883-4412-b8ad-4d6d36471a0c",
            "fb4970d3-3045-049b-46de-3b058eead889",
            "174b4265-e813-4427-a8ce-5eb17b327888",
            "aa9dfbd3-74ae-4543-8203-84ad1a1f4d7f",
            "e862f372-c2d7-4883-bda9-6b9e5f054824",
            "64349c71-c662-40b3-bee8-dfe6f105596d","2d171253-9f2a-430f-8569-d7fad88c8cd0",
            "1b65813d-c4b8-46f7-b596-d5b17d13fda2","01cb2514-9287-458b-b07b-32b887f2435e",
            "37fdc20c-5cd3-4a7f-9518-51b3f85015e8","912c3281-b785-4831-b592-84200ea131d1",
            "b633387a-cbcf-4d0e-a80c-9b54977789b8","7048f754-e328-48b2-90f5-4739b05c3ea3",
            "51f44aad-7af6-454a-b1c2-955641d0c4b4","f86c0e3a-5807-46a4-8e77-e2c6388b05b6",
            "bc1e2f70-3b58-41cd-8406-aaa550482972","e328ca81-0c4e-4109-9208-cc4dfcc018b0",
            "0799f357-efd9-45b0-acf5-1e2915bcdc1a","085575a4-4305-4a2e-a504-e08b97a55a9b",
            "73fe0ef2-7c66-23b9-44e8-8880912e38b1","5571000f-4eb9-5ab3-4d43-33f49858935e",
            "60eef35e-7e67-6dbf-4921-cb96750dcc8a","f680b91c-17eb-3ab9-44db-1214bec56ce9",
            "5c2d914c-1b89-66a7-45d9-3e4600dcc6aa","e413f8b3-18cc-4dee-a13b-949f6f4ca05a",
            "3c616449-2605-4755-b285-578b0f4090bb","fd023367-a07a-4439-a1c2-639003845b9c",
            "3010fdf4-5eb7-1fa6-4521-29bf4e6cc369","c223dfcc-5a2e-4f01-8543-d7ed4411607b",
            "b55d3df1-b3f3-447f-9f65-bc399cb021c5","d7d4b580-6f28-4d23-8db3-68b0c9094489",
            "5c63e4d6-b0cd-4180-b1b6-f99f191a01ae","7b94cd9c-ce01-4041-986b-cfa85feb3b5b",
            "2333e5f5-6afe-4ded-b221-44e133a6dd4b","47c7c72e-101d-4902-95ba-45229f0d92c1",
            "15f2be97-1c97-4959-8a7e-5c2655f52f61","02b3ce0c-c498-4d1c-a335-d52a47fa73fc",
            "8e54f177-d627-4229-8812-9e12a3830ddc","daf47425-b040-470a-9957-c1516dc22006",
            "a9ef8494-d7ec-41d7-900d-079a7a91587e","d4dc5e97-6189-4c3f-a6ed-8626ca67d6f6",
            "7572d67d-b4a9-4b7c-81a0-d43da0fb8fd3","cf5e5121-9695-46b0-9d00-983a9f9792fb",
            "bdd735f5-4be6-4ab1-830e-7782cfe56bcb","b236bc5a-9a06-4676-b458-e6e6c1c65f59",
            "3b4de188-306d-46b7-bc5a-9ea555cf318f","41d7120f-f11c-4342-8474-380b2ad80fd0",
            "6ec80695-3b7c-47d5-86b8-573975e7b2b1","2e85c9df-7958-42e2-87f6-f4de430d491d",
            "6baba687-dd09-4bf5-bbb4-409be8a5b882","f323e851-2465-41dc-94d7-a56290b1d750",
            "9404daf1-384e-f0a7-42ad-6341f4acc299","123f20ad-a52c-49a6-817a-2e390716b449",
            "9092ba53-93c2-72b7-4b19-0f2f1f194c3c","98eb4648-18d7-47d0-89c8-5bebb60f110e",
            "7af41ca9-a343-4e65-afaa-a44666457896","2ef64e39-c4fe-44f0-8f26-7968c77908db",
            };

        private bool opened = true;
        private StringCollection openedDockers = new StringCollection();


        public ControlUI(object app)
        {
            InitializeComponent();
            try
            {
                this.corelApp = app as corel.Application;
                stylesController = new Styles.StylesController(this.Resources, this.corelApp);
            }
            catch
            {
                global::System.Windows.MessageBox.Show("VGCore Erro");
            }
            btn_Command.Click += (s, e) => { openClose((bool)(s as CheckBox).IsChecked); };
        }

        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            stylesController.LoadThemeFromPreference();
            toggleCheckBox(Properties.Settings.Default.IsChecked);
            openClose(this.opened);
        }
        private void toggleCheckBox(bool open)
        {
            this.opened = open;
            btn_Command.IsChecked = open;
            Properties.Settings.Default.IsChecked = open;
            Properties.Settings.Default.Save();
        }
        private void GetOpenedDocker()
        {
            for (int i = 0; i < dockersGuids.Length; i++)
            {
                if (onScreen(dockersGuids[i])&&!openedDockers.Contains(dockersGuids[i]))
                {
                    openedDockers.Add(dockersGuids[i]);
                }
                if (!onScreen(dockersGuids[i]) && openedDockers.Contains(dockersGuids[i]))
                {
                    openedDockers.Remove(dockersGuids[i]);
                }
            }
            Properties.Settings.Default.OpenedList = openedDockers;
            Properties.Settings.Default.Save();
        }
        
        private void openClose(bool open)
        {
            try
            {
                toggleCheckBox(open);
              
                if (opened)
                {
                    for (int i = 0; i < openedDockers.Count; i++)
                    {
                        corelApp.FrameWork.ShowDocker(openedDockers[i]);
                        iDebug("Showed:{0}", openedDockers[i]);
                    }
                }
                else
                {
                    GetOpenedDocker();
                    for (int i = 0; i < openedDockers.Count; i++)
                    {
                        iDebug("Hided:{0}", openedDockers[i]);
                        corelApp.FrameWork.HideDocker(openedDockers[i]);
                        
                    }
                }
            }
            catch (Exception e)
            {
                corelApp.MsgShow(e.Message, "Erro");
            }
        }
        private void iDebug(string textFormat,string captionItem)
        {
            try
            {
                Debug.WriteLine(string.Format(textFormat, corelApp.FrameWork.Automation.GetCaptionText(captionItem)));
            }
            catch(Exception e)
            {
                Debug.WriteLine(string.Format("Erro:{0}",captionItem));
            }
            
        }
        private bool onScreen(string guid)
        {
            int x = 0;
            int y = 0;
            int w = 0;
            int h = 0;

            bool visible = corelApp.FrameWork.IsDockerVisible(guid);
            if (!visible)
                return visible;

            corelApp.FrameWork.Automation.GetItemScreenRect(guid, guid, out x, out y, out w, out h);
            if (w == 0)
                return false;
            return true;
        }
    }
}
