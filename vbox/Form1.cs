using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using VirtualBox;

namespace vbox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            VirtualBox.IVirtualBox vb = new VirtualBox.VirtualBoxClass();
            var mc = vb.Machines;

            foreach (var item in mc)
            {
                var machine = (item as IMachine);

                var nm = machine.Name;

                comboBox1.Items.Add(nm);
                if (machine.State == MachineState.MachineState_Running)
                {
                    comboBox1.SelectedItem = nm;
                }
            }

            bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);

            pictureBox1.Image = bmp;
            pictureBox1.SizeChanged += PictureBox1_SizeChanged;

            gr = Graphics.FromImage(bmp);
            MouseWheel += Form1_MouseWheel;
        }

        #region fields
        string currentMachineName = "";
        Graphics gr;
        Bitmap bmp;
        Thread th;
        VirtualBox.Session session;
        IMachine machine = null;
        object lockobj = new object();
        int msbtnstate = 0;
        #endregion

        private void Form1_MouseWheel(object sender, MouseEventArgs e)
        {
            if (session == null) return;
            var d = session.Console.Mouse;
            d.PutMouseEvent(0, 0, -e.Delta, 0, 0);
        }

        private void PictureBox1_SizeChanged(object sender, EventArgs e)
        {
            bmp = new Bitmap(pictureBox1.Width, pictureBox1.Height);
            pictureBox1.Image = bmp;
            gr = Graphics.FromImage(bmp);
        }

        public Bitmap GetScreenshot()
        {
            //if(session.State==VirtualBox.SessionState.SessionState_Locked)
            lock (lockobj)
            {
                var d = session.Console.Display;
                uint aw, ah, abp;
                int xo, yo;
                VirtualBox.GuestMonitorStatus ms;
                d.GetScreenResolution(0, out aw, out ah, out abp, out xo, out yo, out ms);
                Stopwatch sw = Stopwatch.StartNew();
                var arr = d.TakeScreenShotToArray(0, aw, ah, VirtualBox.BitmapFormat.BitmapFormat_BGRA);
                sw.Stop();

                var ms1 = sw.ElapsedMilliseconds;

                int columns = (int)aw;
                int rows = (int)ah;
                int stride = columns;
                //byte[] newbytes = PadLines(arr, rows, columns);

                Bitmap im = new Bitmap(columns, rows, stride * 4,
                         PixelFormat.Format32bppArgb,
                         Marshal.UnsafeAddrOfPinnedArrayElement(arr.Clone() as Array, 0));

                im = im.Clone(new Rectangle(0, 0, im.Width, im.Height), im.PixelFormat);
                gr.DrawImage(im, new Rectangle(0, 0, bmp.Width, bmp.Height), new Rectangle(0, 0, im.Width, im.Height), GraphicsUnit.Pixel);
                pictureBox1.Invalidate();
                return im;
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            var im = GetScreenshot();
            Clipboard.SetImage(im);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentMachineName))
            { MessageBox.Show("choose machine name first"); return; }
            VirtualBox.IVirtualBox vb = new VirtualBox.VirtualBoxClass();
            var mc = vb.FindMachine(currentMachineName);

            session = new VirtualBox.SessionClass();


            //s.UnlockMachine();

            var ret = mc.LaunchVMProcess(session, "gui", null);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            session.UnlockMachine();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentMachineName))
            { MessageBox.Show("choose machine name first"); return; }
            VirtualBox.IVirtualBox vb = new VirtualBox.VirtualBoxClass();
            var mc = vb.FindMachine(currentMachineName);
            if (mc.State != MachineState.MachineState_Running)
            {
                MessageBox.Show("run machine first");
                return;
            }
            session = new VirtualBox.SessionClass();
            //s.UnlockMachine();
            machine = mc;
            mc.LockMachine(session, VirtualBox.LockType.LockType_Shared);
            UpdateSnapshots(mc);

        }


        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (session != null)
            {
                var d = session.Console.Keyboard;
                var mappedChar = (char)(msg.LParam.ToInt64() >> 16);
                d.PutScancode(mappedChar);
                d.PutScancode((int)mappedChar + 0x80);
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            if (session == null) return;
            if (session.State != SessionState.SessionState_Locked) return;
            if (session.Console == null) return;

            if (checkBox3.Checked)
            {
                SendMouse();
            }

            if (checkBox1.Checked)
            {
                if (th == null)
                {
                    //th = new Thread(() =>
                    // {
                    Stopwatch sw = Stopwatch.StartNew();
                    GetScreenshot();
                    sw.Stop();
                    toolStripStatusLabel1.Text = "screenshot: " + sw.ElapsedMilliseconds + "ms";
                    th = null;
                    //   });
                    // th.IsBackground = true;
                    //  th.Start();
                }

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            var d = session.Console;
            d.PowerButton();
        }

        private void PictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (session == null) return;
            if (e.Button == MouseButtons.Left)
            {
                msbtnstate = 1;
            }
            if (e.Button == MouseButtons.Right)
            {
                msbtnstate = 1 << 1;
            }
            SendMouse();
        }

        void SendMouse()
        {
            if (session.State != SessionState.SessionState_Locked) return;
            lock (lockobj)
            {
                var d = session.Console.Display;

                uint aw, ah, abp;
                int xo, yo;
                VirtualBox.GuestMonitorStatus ms;
                d.GetScreenResolution(0, out aw, out ah, out abp, out xo, out yo, out ms);


                var pos = pictureBox1.PointToClient(Cursor.Position);
                var xrel = (float)pos.X / pictureBox1.Width;
                var yrel = (float)pos.Y / pictureBox1.Height;
                xrel *= aw;
                yrel *= ah;

                var m = session.Console.Mouse;


                m.PutMouseEventAbsolute((int)xrel, (int)yrel, 0, 0, msbtnstate);
            }
        }

        private void PictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (session == null) return;
            var d = session.Console.Mouse;
            msbtnstate = 0;
            SendMouse();
        }


        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            currentMachineName = comboBox1.SelectedItem as string;
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            var d = session.Console;
            d.Reset();

        }

        private void PictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            pictureBox1.Focus();
        }

        private void Button6_Click_1(object sender, EventArgs e)
        {

            var d = session.Console.Keyboard;

            d.ReleaseKeys();
        }

        private void RestoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                var snap = listView1.SelectedItems[0].Tag as ISnapshot;

                machine.LockMachine(session, VirtualBox.LockType.LockType_Shared);

                IProgress prog = session.Machine.RestoreSnapshot(snap);
                prog.WaitForCompletion(300000);
                if (prog.Completed == 1)
                {
                    MessageBox.Show("Operation completed");
                }
                else
                {
                    MessageBox.Show("Operation not completed");
                }
                session.UnlockMachine();
            }
        }

        void UpdateSnapshots(IMachine mc)
        {

            List<ISnapshot> snaps = new List<ISnapshot>();
            listView1.Items.Clear();
            if (mc.SnapshotCount > 0)
            {
                var root = mc.FindSnapshot(null);
                //while (root != null)
                {
                    var snap = root as ISnapshot;
                    snaps.Add(snap);
                    listView1.Items.Add(new ListViewItem(new string[] { root.Name, root.TimeStamp + "" }) { Tag = snap });
                    foreach (var item in snap.Children)
                    {
                        snap = item as ISnapshot;
                        snaps.Add(snap);
                        listView1.Items.Add(new ListViewItem(new string[] { snap.Name, snap.TimeStamp + "" }) { Tag = snap });
                    }

                    //root = root.Children;
                }
            }
        }
        private void Button7_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentMachineName))
            { MessageBox.Show("choose machine name first"); return; }
            VirtualBox.IVirtualBox vb = new VirtualBox.VirtualBoxClass();
            var mc = vb.FindMachine(currentMachineName);

            session = new VirtualBox.SessionClass();

            machine = mc;
            UpdateSnapshots(mc);
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            string aId;

            machine.LockMachine(session, VirtualBox.LockType.LockType_Shared);

            IProgress prog = session.Machine.TakeSnapshot("newsnap", "", 10, out aId);
            prog.WaitForCompletion(300000);
            if (prog.Completed == 1)
            {
                MessageBox.Show("Operation completed");
            }
            else
            {
                MessageBox.Show("Operation not completed");
            }
            session.UnlockMachine();


        }

        private void DeleteSnapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                var snap = listView1.SelectedItems[0].Tag as ISnapshot;

                machine.LockMachine(session, VirtualBox.LockType.LockType_Shared);

                IProgress prog = session.Machine.DeleteSnapshot(snap.Id);
                prog.WaitForCompletion(300000);
                if (prog.Completed == 1)
                {
                    MessageBox.Show("Operation completed");
                }
                else
                {
                    MessageBox.Show("Operation not completed");
                }
                session.UnlockMachine();
            }
        }

        private void Button9_Click(object sender, EventArgs e)
        {
            if (session == null) return;
            
            //session.Machine.GetNetworkAdapter
            var n = session.Machine.GetNetworkAdapter(0) as INetworkAdapter;
            
            var nn = n.NATNetwork;
            Array ar1;
            n.GetProperties("", out ar1);
            var res = n.GetProperty(ar1.GetValue(0) as string);
            n.SetProperty(ar1.GetValue(0) as string, "False");
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            var prcs = Process.GetProcesses().ToArray();
            //enum all windows
            var pc = prcs.FirstOrDefault(z => z.ProcessName.ToLower().Contains(currentMachineName));
            if (pc != null)
            {
                //send minmize

            }
        }
    }
}
