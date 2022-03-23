using System;
using System.Media;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Net;
using System.Net.Sockets;
using System.Collections;
using System.Threading;
using System.IO.Ports;

using System.Text.RegularExpressions;

using System.IO;
using System.Diagnostics;
using OpenCvSharp;
using OpenCvSharp.Extensions;

using NsExcel = Microsoft.Office.Interop.Excel;



namespace Eyebox_Measurement_220216
{



    public partial class Form1 : Form
    {

        Socket sock;

        byte dummy = 0xff;
        byte stx = 0x02;
        byte etx = 0x03;
        byte ACK = 0x06;
        byte nak = 0x15;
        byte rst = 0x12;

        string x_position, z_position, w_position;

        string spd_str = "0000";
        string ltn_stra = "00000.0000";

        string status = "";


        char[] x_pos = new char[10];
        char[] z_pos = new char[10];
        char[] w_pos = new char[10];

        double x_start, z_start;
        double x_stage_point, z_stage_point, w_stage_point;
        decimal x_complete, z_complete, w_complete;


        bool continue_signal = true;




        Mat frame_l = new Mat();
        Mat frame_r = new Mat();

        bool shared_sending, lumi_updated, img_updated, img_writing, show_cam;

        bool cam_opened = false;
        bool stage_opened = false;
        bool lumi_opened = false;


        //int PATTERN_ROW = 5; //가로
        //int PATTERN_COL = 2; //세로
        //double PATTERN_PITCH = 1; //cm


        Mat test_img_gray = new Mat();

        VideoCapture video_cap_right = new VideoCapture(0);
        VideoCapture video_cap_left = new VideoCapture(1);

        //OpenCvSharp.Size src_size = new OpenCvSharp.Size(2048, 1536);
        OpenCvSharp.Size src_size = new OpenCvSharp.Size(640, 480);


        bool img_pixel_get_signal = true;

        List<Tuple<string, double, double>> pixel_ever_result_right = new List<Tuple<string, double, double>>();
        List<Tuple<string, double, double>> pixel_ever_result_left = new List<Tuple<string, double, double>>();

        //List<Tuple<int, int>> tuple_test = new List<Tuple<int, int>>();
        //int[] pixel_ever_result = new int[];

        SerialPort _serialPort_luminance = new SerialPort();

        double save_label = 0;

        static double save_label_start = 0;
        static double save_label_end = 0;
        static double measure_center_point = 0;


        string[] comlist = System.IO.Ports.SerialPort.GetPortNames();





        public static byte[] Combine(byte[] first, byte[] second) //byte 결합하는 함수에 관한 부분
        {
            return first.Concat(second).ToArray();
        }

        public byte lrc_cal(byte[] data)  //명령어 LRC 계산하는 부분
        {
            //byte XOR 연산
            byte lrc = dummy;

            for (int n = 0; n < data.Length; n++)
            {
                lrc = (byte)(lrc ^ data[n]);
            }

            if (lrc == 0)
            {
                lrc = etx;
            }
            return lrc;
        }




        public Form1()
        {
            InitializeComponent();


            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

        }

        private void send_basic_func(byte[] msg)
        {
            byte[] bytes = new byte[50];

            byte[] ack = new byte[] { ACK };
            byte[] header = new byte[] { stx, dummy };

            msg = Combine(header, msg);

            if (sock.Available > 0) // here we clean up the current queue
            {
                sock.Receive(bytes);
            }

            sock.Send(msg);

            while (sock.Available == 0) // wait for the controller response
            {
                Thread.Sleep(100);
            }

            sock.Receive(bytes); // after receiving the data, we should check the LRC if possible
                                 //string status = Encoding.UTF8.GetString(bytes);
                                 // bytes[3] -> 채널0 정보 포함 
                                 // 10진수(180) -> 2진수 변환
                                 // 1 0 1 1 0 1 0 0 (2진수)
                                 // 7 6 5 4 3 2 1 0
                                 // Bit5: Servo On / Bit4: Origin / Bit3: Alarm 
                                 // Bit2: Ready / Bit1: In Position / Bit0: Run


            if (bytes.Contains<byte>(nak) || bytes.Contains<byte>(rst) == true)
            {
                sock.Send(msg);
            }
            else
            {
                sock.Send(ack);
            }
            //receive_data = receive_lrc_cal(bytes);
            //status = Encoding.UTF8.GetString(bytes);
        }


        private void send_function(byte[] msg)
        {
            byte[] bytes = new byte[50];

            byte[] ack = new byte[] { ACK };
            byte[] header = new byte[] { stx, dummy };

            msg = Combine(header, msg);

            if (sock.Available > 0) // here we clean up the current queue
            {
                sock.Receive(bytes);
            }

            sock.Send(msg);

            while (sock.Available == 0) // wait for the controller response
            {
                Thread.Sleep(100);
            }

            sock.Receive(bytes); // after receiving the data, we should check the LRC if possible
                                 //string status = Encoding.UTF8.GetString(bytes);

            if (bytes.Contains<byte>(nak) || bytes.Contains<byte>(rst) == true)
            {
                sock.Send(msg);
            }
            else
            {
                sock.Send(ack);
            }

            status = Encoding.UTF8.GetString(bytes);
        }



        public byte[] speed(string channel, string spd)
        {
            byte[] command = Encoding.UTF8.GetBytes("CB");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] speed_set = Encoding.UTF8.GetBytes(spd);
            byte[] make_msg = Combine(command, channel_ba);
            make_msg = Combine(make_msg, speed_set);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }








        private void robo_con_Click(object sender, EventArgs e)
        {
            sock = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            IPAddress ip = IPAddress.Parse("192.168.1.203");//인자값 : 서버측 IP         
            IPEndPoint endPoint = new IPEndPoint(ip, 20000);//인자값 : IPAddress,포트번호

            while (sock.Connected == false)
            {
                sock.Connect(endPoint);
            }

            if (sock.Connected == true)
            {
                textBox1.Text = ("Stage Connected");
                robo_con.Enabled = false;
            }
        }

        public byte[] move_zero(string channel) // 원점 이동 부분
        {
            byte[] command = Encoding.UTF8.GetBytes("BA");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] make_msg = Combine(command, channel_ba);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }



        private void origin_Click(object sender, EventArgs e)
        {
            var comm = move_zero("0");
            send_function(comm);
        }

        public byte[] servo_on(string channel)
        {
            byte[] command = Encoding.UTF8.GetBytes("DB");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] data_type = Encoding.UTF8.GetBytes("1");
            byte[] make_msg = Combine(command, channel_ba);
            make_msg = Combine(make_msg, data_type);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }


        private void sv_on_Click(object sender, EventArgs e)
        {
            var comm = servo_on("0");
            send_function(comm);
        }


        public byte[] servo_off(string channel)
        {
            byte[] command = Encoding.UTF8.GetBytes("DB");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] data_type = Encoding.UTF8.GetBytes("0");
            byte[] make_msg = Combine(command, channel_ba);
            make_msg = Combine(make_msg, data_type);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }



        private void sv_off_Click(object sender, EventArgs e)
        {
            var comm = servo_off("0");
            send_function(comm);
        }

        public byte[] posi_check_robot(string channel) //robot_position chk
        {
            byte[] command = Encoding.UTF8.GetBytes("AC");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] data_type = Encoding.UTF8.GetBytes("2");
            byte[] make_msg = Combine(command, channel_ba);
            make_msg = Combine(make_msg, data_type);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }

        private void send_position(byte[] msg) //좌표 수신
        {

            byte[] bytes = new byte[50];

            byte[] ack = new byte[] { ACK };
            byte[] header = new byte[] { stx, dummy };

            msg = Combine(header, msg);

            if (sock.Available > 0) // here we clean up the current queue
            {
                sock.Receive(bytes);
                //sock1.Receive(bytes1);
            }

            sock.Send(msg);

            while (sock.Available == 0) // wait for the controller response
            {
                Thread.Sleep(100);
            }

            sock.Receive(bytes); // after receiving the data, we should check the LRC if possible

            if (bytes.Contains<byte>(nak) || bytes.Contains<byte>(rst) == true)
            {
                sock.Send(msg);
            }
            else
            {
                sock.Send(ack);
            }

            status = Encoding.UTF8.GetString(bytes);

            status.CopyTo(3, x_pos, 0, 8);
            //status.CopyTo(13, y_pos, 0, 7);
            status.CopyTo(23, z_pos, 0, 8);
            status.CopyTo(33, w_pos, 0, 8);

            //int val = Convert.ToInt32(x_pos[8]);
            //x_pos[8] = Convert.ToChar(48);//Convert.ToChar(val - 1);
            //int x_pos_abs = Convert.ToInt16(x_pos);

            x_position = new string(x_pos);//비교를 위해 초기값 저장
            x_position = x_position.Trim();

            decimal x_pos_abs = Convert.ToDecimal(x_position);
            x_pos_abs = Math.Abs(x_pos_abs);
            x_position = Convert.ToString(x_pos_abs);

            z_position = new string(z_pos);
            z_position = z_position.Trim();

            decimal z_pos_abs = Convert.ToDecimal(z_position);
            z_pos_abs = Math.Abs(z_pos_abs);
            z_position = Convert.ToString(z_pos_abs);

            msg.Initialize();
            bytes.Initialize();
        }

        public void SetText()
        {
            lb_x_position.Text = x_position + " mm";
            lb_z_position.Text = z_position + " mm";
        }


        private void posi_chk_Click(object sender, EventArgs e)
        {
            var comm = posi_check_robot("0");
            send_position(comm);
            this.Invoke(new Action(SetText));
        }


        public byte[] move_axis_all_z(string channel, double x_axis_point, double z_axis_point)
        {
            byte[] command = Encoding.UTF8.GetBytes("BC");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] motion_type = Encoding.UTF8.GetBytes("0");
            byte[] xy_type = Encoding.UTF8.GetBytes("1");

            decimal null_byte = 0;
            string null_st_fn = null_byte.ToString(ltn_stra);

            decimal x_st = Convert.ToDecimal(x_axis_point);
            string x_st_fn = x_st.ToString(ltn_stra);
            byte[] location_x = Encoding.UTF8.GetBytes(x_st_fn);

            string y_st_fn = ltn_stra.ToString();
            byte[] location_y = Encoding.UTF8.GetBytes(y_st_fn);

            decimal z_st = Convert.ToDecimal(z_axis_point);
            string z_st_fn = z_st.ToString(ltn_stra);
            byte[] location_z1 = Encoding.UTF8.GetBytes(z_st_fn);

            decimal w_st = Convert.ToDecimal(z_axis_point);
            string w_st_fn = z_st.ToString(ltn_stra);
            byte[] location_w1 = Encoding.UTF8.GetBytes(z_st_fn);

            //byte[] location_null_z1 = Encoding.UTF8.GetBytes(null_st_fn);

            //byte[] location_z2 = Encoding.UTF8.GetBytes(z_st_fn);
            //byte[] location_null_z2 = Encoding.UTF8.GetBytes(null_st_fn);

            //byte[] xy_location_final = Combine(xy_location1, xy_location2);
            //byte[] xy_location_final1 = Combine(xy_location2, xy_location3);
            //byte[] xy_location_final2 = Combine(xy_location3, xy_location4);

            //byte[] xy_location_final = Combine(xy_location1, xy_location_null);
            //xy_location_final = Combine(xy_location_final, xy_location2);
            //xy_location_final = Combine(xy_location_final, xy_location_null2);

            byte[] make_msg = Combine(command, channel_ba);
            make_msg = Combine(make_msg, motion_type);
            make_msg = Combine(make_msg, xy_type);

            make_msg = Combine(make_msg, location_x);
            make_msg = Combine(make_msg, location_y);
            make_msg = Combine(make_msg, location_z1);
            make_msg = Combine(make_msg, location_z1);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }


        private async void setup_point_Click(object sender, EventArgs e)
        {
            this.Invoke(new Action(delegate ()
            {
                textBox3.Clear();
                richTextBox1.Clear();
                progressBar1.Value = 0;


                if (graph_draw)
                { 
                    chart1.Series["Eyebox"].Points.Clear();

                    graph_draw = false;
                }

            }));

            decimal speed_2 = spd_value_start.Value;
            string spd = speed_2.ToString(spd_str);
            var comm = speed("0", spd);
            send_function(comm);

            x_stage_point = Convert.ToDouble(x_srt.Value);
            z_stage_point = Convert.ToDouble(z_srt.Value);
            w_stage_point = Convert.ToDouble(z_srt.Value);

            double x_set = Convert.ToDouble(x_position);
            double z_set = Convert.ToDouble(z_position);
            double w_set = Convert.ToDouble(w_position);

            Thread.Sleep(300);

            x_start = x_stage_point;
            z_start = z_stage_point;

            double x_setup_db = CustomRound(RoundType.Truncate, x_stage_point, 3);
            double z_setup_db = CustomRound(RoundType.Truncate, z_stage_point, 3);

            var task_setup = Task.Run(() =>
            {
                var comm_move_xz = move_axis_all_z("0", x_stage_point, z_stage_point);
                send_function(comm_move_xz);

                Thread.Sleep(1000);


                while (true)
                {
                    var comm_posi = posi_check_robot("0");
                    send_position(comm_posi);

                    x_complete = Convert.ToDecimal(x_position);
                    z_complete = Convert.ToDecimal(z_position);

                    if (x_complete == Convert.ToDecimal(x_setup_db) && z_complete == Convert.ToDecimal(z_setup_db))
                    // y축 좌표 비교 이동 완료
                    {
                        this.Invoke(new Action(setup_ok_print));

                        img_pixel_get_signal = false;
                        
                        break;
                    }
                    else
                    {
                        var comm_move_xz_re = move_axis_all_z("0", x_stage_point, z_stage_point);
                        send_function(comm_move_xz_re);
                    }
                }
            });
            await task_setup;

        }



        public void setup_ok_print()
        {
            textBox3.ResetText();
            textBox3.Text = "setup";

            MessageBox.Show("Start Point OK");

        }

        public void setup_ok_lumi()
        {
            //textBox3.ResetText();
            richTextBox2.AppendText("setup");

            //MessageBox.Show("Start Point OK");

        }



        static private double CustomRound(RoundType roundType, double value, int digit = 1)
        {
            double dReturn = 0;

            // 지정 자릿수의 올림,반올림, 버림을 계산하기 위한 중간 계산
            double digitCal = Math.Pow(10, digit) / 10;

            switch (roundType)
            {
                case RoundType.Ceiling:
                    dReturn = Math.Ceiling(value * digitCal) / digitCal;
                    break;
                case RoundType.Round:
                    dReturn = Math.Round(value * digitCal) / digitCal;
                    break;
                case RoundType.Truncate:
                    dReturn = Math.Truncate(value * digitCal) / digitCal;
                    break;
            }
            return dReturn;
        }

        private enum RoundType
        {
            Ceiling,
            Round,
            Truncate
        }

        public byte[] jog_start(string channel, string axis, string pm)
        {
            byte[] command = Encoding.UTF8.GetBytes("BE");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] axis_type = Encoding.UTF8.GetBytes(axis);
            byte[] direc_type = Encoding.UTF8.GetBytes(pm);
            byte[] motion_type = Encoding.UTF8.GetBytes("0");
            byte[] make_msg = Combine(command, channel_ba);
            make_msg = Combine(make_msg, axis_type);
            make_msg = Combine(make_msg, direc_type);
            make_msg = Combine(make_msg, motion_type);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }



        public byte[] jog_continue(string channel)
        {
            byte[] command = Encoding.UTF8.GetBytes("BF");
            byte[] channel_ba = Encoding.UTF8.GetBytes(channel);
            byte[] make_msg = Combine(command, channel_ba);

            byte lrc = lrc_cal(make_msg);

            byte[] etx_ba = new byte[] { etx };
            byte[] lrc_ba = new byte[] { lrc };

            make_msg = Combine(make_msg, etx_ba);
            make_msg = Combine(make_msg, lrc_ba);

            return make_msg;
        }





        private async void x_left_MouseDown(object sender, MouseEventArgs e)
        {
            continue_signal = true;

            decimal speed_1 = spd_value.Value;
            string spd = speed_1.ToString(spd_str);

            var comm_spd = speed("0", spd);
            send_basic_func(comm_spd);

            Thread.Sleep(100);

            var comm = jog_start("0", "0", "0");
            send_basic_func(comm);

            Thread.Sleep(100);

            var continue_task = Task.Run(() =>
            {
                while (continue_signal)
                {
                    var continue_comm = jog_continue("0");
                    send_basic_func(continue_comm);

                    Thread.Sleep(100);

                    if (continue_signal == false)
                    {
                        break;
                    }
                }
            });
            await continue_task;
        }

        private void x_left_MouseUp(object sender, MouseEventArgs e)
        {
            continue_signal = false;

        }



        public unsafe async void eyebox_searching()
        {
            
            cam_opened = true;

            Mat frame_right_test = new Mat();
            Mat gray_right = new Mat();


            Mat frame_left_test = new Mat();
            Mat gray_left = new Mat();

            video_cap_right.Set(CaptureProperty.FrameWidth, 640);
            video_cap_right.Set(CaptureProperty.FrameHeight, 480);

            video_cap_left.Set(CaptureProperty.FrameWidth, 640);
            video_cap_left.Set(CaptureProperty.FrameHeight, 480);


            Stopwatch watch = new Stopwatch();
            int sum = 0;

            
            double x_pitch_save = Convert.ToDouble(x_pitch_val.Value);


            //video_frame.Set(CaptureProperty.FrameWidth, 2048);
            //video_frame.Set(CaptureProperty.FrameHeight, 1536);

            //Cv2.NamedWindow("Left", WindowMode.Normal);
            //Cv2.ResizeWindow("Left", 640, 480);

            img_writing = false;
            show_cam = true;

            double b_sum_right = 0;
            double result_right = 0;

            double b_sum_left = 0;
            double result_left = 0;

            var img_gray_task = Task.Run(() =>
            {
                watch.Start();

                for (; ; )
                {

                    if (!img_writing)
                    {
                        video_cap_right.Read(frame_right_test);
                        //Cv2.Flip(frame_l_test, frame_l_test, FlipMode.XY);
                        Cv2.CvtColor(frame_right_test, gray_right, ColorConversionCodes.BGR2GRAY);
                        Cv2.ImWrite("./test_gray_right.jpg", gray_right);

                        video_cap_left.Read(frame_left_test);
                        //Cv2.Flip(frame_l_test, frame_l_test, FlipMode.XY);
                        Cv2.CvtColor(frame_left_test, gray_left, ColorConversionCodes.BGR2GRAY);
                        Cv2.ImWrite("./test_gray_left.jpg", gray_left);

                        //Mat test_img_gray1 = Cv2.ImRead("test_frame.jpg", ImreadModes.Grayscale);
                        //Cv2.ImReadMulti()
                        //Cv2.ImWrite("test_gray.jpg", test_img_gray1);
                        //test_img_gray1.Dispose();


                        for (int row = 0; row < gray_right.Rows; row++)
                        {
                            for (int col = 0; col < gray_right.Cols; col++)
                            {
                                byte b_right;//r, g, b;
                                byte b_left;

                                // byte *data = (byte *)image.Data.ToPointer();
                                byte* data_right = (byte*)gray_right.DataPointer;
                                byte* data_left = (byte*)gray_left.DataPointer;

                                b_right = data_right[row * gray_right.Step() + col * gray_right.ElemSize() + 0];
                                b_sum_right += b_right;

                                //BitConverter.ToInt32(b, 0);
                                b_left = data_left[row * gray_right.Step() + col * gray_right.ElemSize() + 0];
                                b_sum_left += b_left;

                                //textBox4.Invoke((MethodInvoker)delegate ()
                                //{
                                //    textBox4.AppendText(b.ToString() + "\r" + ", ");
                                //});
                                //Textbox4.

                                //g = data[row * frame_l.Step() + col * frame_l.ElemSize() + 1];
                                //r = data[row * frame_l.Step() + col * frame_l.ElemSize() + 2];

                                //data[row * frame_l.Step() + col * frame_l.ElemSize() + 0] = (byte)(255 - b);
                                //data[row * frame_l.Step() + col * frame_l.ElemSize() + 1] = (byte)(255 - g);
                                //data[row * frame_l.Step() + col * frame_l.ElemSize() + 2] = (byte)(255 - r);

                                //System.IntPtr b = img_data[row * frame_l.Cols * 3 + col * 3];
                                //System.IntPtr g = img_data[row * frame_l.Cols * 3 + col * 3 + 1];
                                //System.IntPtr r = img_data[row * frame_l.Cols * 3 + col * 3 + 2];
                            }
                        }


                    }


                    img_updated = true;

                    if (img_updated)
                    {
                        img_updated = false;
                        break;
                    }
                }

                watch.Stop();
                long a = watch.ElapsedMilliseconds; //1초 소비

                result_right = b_sum_right / (gray_right.Rows * gray_right.Cols);
                
                result_left = b_sum_left / (gray_right.Rows * gray_right.Cols);

                if(x_stage_point <= measure_center_point)
                {
                    pixel_ever_result_right.Add(Tuple.Create("right value", save_label, result_right));

                }
                else 
                {
                    pixel_ever_result_left.Add(Tuple.Create("left value", save_label, result_left));


                }

                save_label += x_pitch_save;

                /*
                foreach (var item in pixel_ever_result)
                ichTextBox1.Invoke((MethodInvoker)delegate ()
                {
                    richTextBox1.AppendText(item.Item1.ToString() + ", " + item.Item2.ToString());
                    richTextBox1.AppendText("\r\n");
                    richTextBox1.ScrollToCaret();
                });
                */


                result_right = 0;
                b_sum_right = 0;

                result_left = 0;
                b_sum_left = 0;

            });
            //await img_gray_task;
            img_gray_task.Wait();

            frame_right_test.Dispose();
            gray_right.Dispose();

            frame_left_test.Dispose();
            gray_left.Dispose();

            //Cv2.Flip(frame_r, frame_r, FlipMode.XY);
            /*
            while (Cv2.WaitKey(33) != 'q')
            {

                //video_L.Read(frame1);
                Cv2.Circle(frame_l, 0, 0, 0, Scalar.Yellow, 10, LineTypes.AntiAlias);

                Cv2.ImShow("Left", frame_l);

                //video_R.Read(frame2);
                //Cv2.ImShow("Right", frame2);
            }
            */
            //Cv2.Ellipse(frame1, );

            //frame_l.Dispose();
            // frame_r.Dispose();

            //video_L.Release();
            //video_R.Release();

            //Cv2.DestroyAllWindows();
            cam_opened = false;


        }


        private async void data_search_Click(object sender, EventArgs e)
        {
            img_pixel_get_signal = true;


            var img_gray_task = Task.Run(() =>
            {
                while(img_pixel_get_signal)
                {
                    eyebox_searching();

                }

                //while (continue_signal)
                //{
                //    img_convert();

                //if (continue_signal == false)
                //{
                //    break;
                //}
                //}
            });
            await img_gray_task;

        }













     

       

        private async void x_right_MouseDown(object sender, MouseEventArgs e)
        {           


            continue_signal = true;


            //decimal speed_2 = spd_value_start.Value;
            //string spd = speed_2.ToString(spd_str);
            //var comm = speed("0", spd);
            //send_function(comm);

            decimal speed_1 = spd_value.Value;
            string spd = speed_1.ToString(spd_str);
            var comm_spd = speed("0", spd);
            send_basic_func(comm_spd);

            Thread.Sleep(100);

            var comm = jog_start("0", "0", "1");
            send_basic_func(comm);

            Thread.Sleep(100);                         

            var continue_task = Task.Run(() =>
            {
                while (continue_signal)
                {
                    var continue_comm = jog_continue("0");
                    send_basic_func(continue_comm);

                    Thread.Sleep(100);

                //img_convert();

                    if (continue_signal == false)
                    {
                            break;
                    }
                }
            });
            await continue_task;


        }

        
        private void x_right_MouseUp(object sender, MouseEventArgs e)
        {
            continue_signal = false;

        }

     
        private async void z_up_MouseDown(object sender, MouseEventArgs e)
        {
            continue_signal = true;

            decimal speed_1 = spd_value.Value;
            string spd = speed_1.ToString(spd_str);

            var comm_spd = speed("0", spd);
            send_basic_func(comm_spd);

            Thread.Sleep(100);

            var comm = jog_start("0", "2", "1");
            send_basic_func(comm);

            Thread.Sleep(100);

            var continue_task = Task.Run(() =>
            {
                while (continue_signal)
                {
                    var continue_comm = jog_continue("0");
                    send_basic_func(continue_comm);

                    Thread.Sleep(100);

                    if (continue_signal == false)
                    {
                        break;
                    }
                }
            });
            await continue_task;
        }

        private void z_up_MouseUp(object sender, MouseEventArgs e)
        {
            continue_signal = false;

        }

        private async void z_down_MouseDown(object sender, MouseEventArgs e)
        {
            continue_signal = true;

            decimal speed_1 = spd_value.Value;
            string spd = speed_1.ToString(spd_str);

            var comm_spd = speed("0", spd);
            send_basic_func(comm_spd);

            Thread.Sleep(100);

            var comm = jog_start("0", "2", "0");
            send_basic_func(comm);

            Thread.Sleep(100);

            var continue_task = Task.Run(() =>
            {
                while (continue_signal)
                {
                    var continue_comm = jog_continue("0");
                    send_basic_func(continue_comm);

                    Thread.Sleep(100);

                    if (continue_signal == false)
                    {
                        break;
                    }
                }
            });
            await continue_task;
        }

        private void z_down_MouseUp(object sender, MouseEventArgs e)
        {
            continue_signal = false;

        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }

        private void lumi_combobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            comlist[0] = lumi_combobox.Text;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {



        }


        private void Form1_Load(object sender, EventArgs e)
        {

            if (comlist.Length > 0)
            {
                lumi_combobox.Items.AddRange(comlist);
                //제일 처음에 위치한 녀석을 선택
                lumi_combobox.SelectedIndex = 0;



            }
        }



        







        private void cam_ready_Click(object sender, EventArgs e)
        {

        }

        private void cam_open_Click(object sender, EventArgs e)
        {
            // OpenCvSharp.

            // 카메라 0, 1번
            // 

            VideoCapture video_cap_right = new VideoCapture(0); //0번 right
            VideoCapture video_cap_left = new VideoCapture(1);  //1번 left

            Mat frame_right = new Mat();
            Mat frame_left = new Mat();


            //VideoCapture video_R = new VideoCapture(1);
            cam_opened = true;

            video_cap_right.Set(CaptureProperty.FrameWidth, 2048);
            video_cap_right.Set(CaptureProperty.FrameHeight, 1536);
            video_cap_left.Set(CaptureProperty.FrameWidth, 2048);
            video_cap_left.Set(CaptureProperty.FrameHeight, 1536);

            Cv2.NamedWindow("Right", WindowMode.Normal);
            Cv2.ResizeWindow("Right", 640, 480);
            Cv2.NamedWindow("Left", WindowMode.Normal);
            Cv2.ResizeWindow("Left", 640, 480);
            
            
            img_writing = false;
            show_cam = true;


            for (; ; )
            {
                if (!img_writing)
                {
                    video_cap_right.Read(frame_right);
                    video_cap_left.Read(frame_left);


                    //Cv2.Flip(frame_l, frame_l, FlipMode.XY);
                    //Cv2.Flip(frame_r, frame_r, FlipMode.XY);
                    img_updated = true;
                }

                Cv2.ImShow("Right", frame_right);
                Cv2.ImShow("Left", frame_left);

                try
                {
                    if (Cv2.WaitKey(10) >= 0 || !show_cam)
                        break;
                }
                catch (Exception ex) { }
            }

            /*
            while (Cv2.WaitKey(33) != 'q')
            {

                //video_L.Read(frame1);
                Cv2.Circle(frame_l, 0, 0, 0, Scalar.Yellow, 10, LineTypes.AntiAlias);

                Cv2.ImShow("Left", frame_l);

                //video_R.Read(frame2);
                //Cv2.ImShow("Right", frame2);

            }
            */
            //Cv2.Ellipse(frame1, );

            frame_right.Dispose();
            frame_left.Dispose();

            video_cap_right.Release();
            video_cap_left.Release();            

            Cv2.DestroyAllWindows();
            cam_opened = false;

        }



        private async void lumi_check_Click(object sender, EventArgs e)
        {
            this.Invoke(new Action(delegate ()
            {
                richTextBox2.Clear();
            }));


            decimal speed_2 = spd_value_start.Value;
            string spd = speed_2.ToString(spd_str);
            var comm = speed("0", spd);
            send_function(comm);

            double x_measure_point = Convert.ToDouble(measure_point_x.Value);
            double z_measure_point = Convert.ToDouble(measure_point_z.Value);
            double w_measure_point = Convert.ToDouble(measure_point_z.Value);

            double x_set = Convert.ToDouble(x_position);
            double z_set = Convert.ToDouble(z_position);
            double w_set = Convert.ToDouble(w_position);

            Thread.Sleep(300);

            double x_start_measure = x_measure_point;
            double z_start_measure = z_measure_point;

            double x_setup_db = CustomRound(RoundType.Truncate, x_start_measure, 3);
            double z_setup_db = CustomRound(RoundType.Truncate, z_start_measure, 3);

            var task_setup = Task.Run(() =>
            {
                var comm_move_xz = move_axis_all_z("0", x_measure_point, z_measure_point);
                send_function(comm_move_xz);

                Thread.Sleep(1000);


                while (true)
                {
                    var comm_posi = posi_check_robot("0");
                    send_position(comm_posi);

                    x_complete = Convert.ToDecimal(x_position);
                    z_complete = Convert.ToDecimal(z_position);

                    if (x_complete == Convert.ToDecimal(x_setup_db) && z_complete == Convert.ToDecimal(z_setup_db))
                    // y축 좌표 비교 이동 완료
                    {
                        this.Invoke(new Action(setup_ok_lumi));

                        img_pixel_get_signal = false;

                        break;
                    }
                    else
                    {
                        var comm_move_xz_re = move_axis_all_z("0", x_measure_point, z_measure_point);
                        send_function(comm_move_xz_re);
                    }
                }
            });
            await task_setup;

            //MessageBox.Show("측정중...");


            if (!_serialPort_luminance.IsOpen)
            {


                _serialPort_luminance.PortName = comlist[0];


                _serialPort_luminance.BaudRate = 4800;
                _serialPort_luminance.Parity = Parity.Even;
                _serialPort_luminance.DataBits = 7;
                _serialPort_luminance.StopBits = StopBits.Two;
                _serialPort_luminance.Handshake = Handshake.RequestToSend;
                _serialPort_luminance.NewLine = "\r\n";

                //장치 관리자 설정 -> 데이터 handshake: 없음

                _serialPort_luminance.Open();

                if (_serialPort_luminance.IsOpen)
                {
                    //Console.WriteLine("lumi connected");
                    //textBox1.AppendText("\r\nLuminance Connected");
                    //MessageBox.Show("connect");

                }
                else
                {

                    MessageBox.Show("not connect");
                    //Console.WriteLine("not connected");
                    //textBox1.AppendText("\r\nLuminance Not Connected");
                    return;
                }
            }



            string[] lumi_value;// = new string[10];
            string lumi_val_str;

            _serialPort_luminance.DiscardInBuffer();

            while (true)
            {
                //Thread.Sleep(100);
                _serialPort_luminance.Write("MES\r\n");

                lumi_val_str = _serialPort_luminance.ReadExisting();
                //lumi_value = serialPort_lumi.ReadExisting();

                Console.WriteLine(lumi_val_str);

                if (lumi_val_str.Contains("OK"))
                    break;
            }

            lumi_value = lumi_val_str.Split(',');

            try
            {
                if (lumi_val_str != "" && !lumi_val_str.Contains("ER"))
                {
                    lumi_value[1] = lumi_value[1].Trim();


                    lumi_value[2] = lumi_value[2].Trim();


                    lumi_value[3] = lumi_value[3].Trim();


                    richTextBox2.Clear();
                    richTextBox2.AppendText("휘도: " + lumi_value[1] + "(cd/m^2)");// + "\r\n색도 x: " + lumi_value[2] + ", y: " + lumi_value[3]);


                }
                else
                {
                    MessageBox.Show("측정 오류, 다시 체크하세요");
                    Array.Clear(lumi_value, 0, lumi_value.Length);
                    return;
                }
            }
            catch (Exception)
            {

                Array.Clear(lumi_value, 0, lumi_value.Length);

            }



        }

        private async void start_job_Click(object sender, EventArgs e)
        {
            //설정 x mm만큼 이동

            save_label = Convert.ToDouble(x_srt.Value);

            save_label_start = save_label;
            save_label_end = Convert.ToDouble(x_end_point.Value);

            measure_center_point = (save_label_start + save_label_end) / 2;

            //this.Invoke(new Action(delegate ()
            //{
            //    textBox3.Clear();
            //}));
            decimal speed_2 = spd_value_start.Value;
            string spd = speed_2.ToString(spd_str);
            var comm = speed("0", spd);
            send_function(comm);

            x_stage_point = Convert.ToDouble(x_srt.Value);
            z_stage_point = Convert.ToDouble(z_srt.Value);
            w_stage_point = Convert.ToDouble(z_srt.Value);

            double x_set = Convert.ToDouble(x_position);
            double z_set = Convert.ToDouble(z_position);
            double w_set = Convert.ToDouble(w_position);

            double x_end = Convert.ToDouble(x_end_point.Value);

            double x_move_pitch = Convert.ToDouble(x_pitch_val.Value);
            int x_job_count = (int)((x_end - x_stage_point) / x_move_pitch);


            if (x_stage_point + (x_move_pitch * x_job_count) > 800.000)
            {
                MessageBox.Show("Stage X Axis Limit");
                return;
            }

            Thread.Sleep(100);

            x_start = x_stage_point;
            z_start = z_stage_point;

            double x_setup_db = CustomRound(RoundType.Truncate, x_stage_point, 3);
            double z_setup_db = CustomRound(RoundType.Truncate, z_stage_point, 3);

            for (int job_count_x = 0; job_count_x < x_job_count + 1; job_count_x++)
            {

                var task_posi_out = Task.Run(() =>
                {

                    this.Invoke(new MethodInvoker(delegate ()
                    {
                        //progressBar1.Value = i + 1;
                        progressBar1.Maximum = x_job_count;
                        progressBar1.Value = job_count_x;

                        richTextBox1.AppendText(x_stage_point.ToString());// + ", " + z_stage_point.ToString());
                        richTextBox1.AppendText("\r\n");
                        richTextBox1.ScrollToCaret();
                        //richTextBox1.AppendText("\r");

                    }));
                    //this.Invoke(new Action(SetText));
                });
                await task_posi_out;

                var eyebox_searching_task = Task.Factory.StartNew(eyebox_searching);
                //await shutter_open_single;
                eyebox_searching_task.Wait();

                x_stage_point += x_move_pitch;

                if(x_stage_point > save_label_end)
                {
                    break;
                }


                var task_setup = Task.Run(() =>
                {
                    var comm_move_xz = move_axis_all_z("0", x_stage_point, z_stage_point);
                    send_function(comm_move_xz);

                    Thread.Sleep(200);


                    while (true)
                    {
                        var comm_posi = posi_check_robot("0");
                        send_position(comm_posi);

                        x_complete = Convert.ToDecimal(x_position);
                        z_complete = Convert.ToDecimal(z_position);

                        if (x_complete == Convert.ToDecimal(x_stage_point) && z_complete == Convert.ToDecimal(z_setup_db))
                        // y축 좌표 비교 이동 완료
                        {
                            //this.Invoke(new Action(setup_ok_print));
                            //img_pixel_get_signal = false;

                            break;
                        }
                        else
                        {
                            var comm_move_xz_re = move_axis_all_z("0", x_stage_point, z_stage_point);
                            send_function(comm_move_xz_re);
                        }
                    }
                });
                await task_setup;










                //if (job_count_x != spot_count_x.Value - 1)
                //{
                //    var task_x_axis = Task.Factory.StartNew(x_job_method);
                //    await task_x_axis;
                //}
                //else
                //{
                //    break;
                //}          

            }


            var data_prograss_task = Task.Run(() =>
            {
                this.Invoke(new MethodInvoker(delegate ()
                {
                    //progressBar1.Value = i + 1;
                    //progressBar1.Maximum = x_job_count;
                    progressBar1.Value = x_job_count;



                }));


                ListToExcel(pixel_ever_result_right, pixel_ever_result_left); //save1
                //ListToExcel(pixel_ever_result_left); //save1


                //var Graph_Task = Task.Run(() => 
                //{

                this.Invoke(new MethodInvoker(delegate ()
                {
                    ListToGraph(pixel_ever_result_right, pixel_ever_result_left);
                    //ListToGraph(pixel_ever_result_left);
                }));
                //});
                //Graph_Task.Wait();


            });
            await data_prograss_task;



            MessageBox.Show("Job Finish");

            //if (job_count_z != spot_count_z.Value - 1)
            //{
            //    var task_z_axis = Task.Factory.StartNew(z_job_method);
            //    await task_z_axis;
            //}
            //else
            //{
            //    break;
            //}


            //task_z_axis.Wait();
        }

        public void ListToExcel(List<Tuple<string, double, double>> list, List<Tuple<string, double, double>> list2)
        {

            //start excel
            NsExcel.Application excapp = new Microsoft.Office.Interop.Excel.Application();

            //if you want to make excel visible           

            //create a blank workbook
            var workbook = excapp.Workbooks.Add(NsExcel.XlWBATemplate.xlWBATWorksheet);

            //or open one - this is no pleasant, but yue're probably interested in the first parameter
            //string workbookPath = @"C:\result\excel_test\test_data.xlsx";
            //var workbook1 = excapp.Workbooks.Open(workbookPath,
            //    0, false, 5, "", "", false, NsExcel.XlPlatform.xlWindows, "",
            //    true, false, 0, true, false, false);

            //Not done yet. You have to work on a specific sheet - note the cast
            //You may not have any sheets at all. Then you have to add one with NsExcel.Worksheet.Add()
            var sheet = (NsExcel.Worksheet)workbook.Sheets[1]; //indexing starts from 1

            //do something usefull: you select now an individual cell
            var range_A = sheet.get_Range("A1", "A1");
            range_A.Value2 = "mm"; //Value2 is not a typo

            var range_B = sheet.get_Range("B1", "B1");
            range_B.Value2 = "RIGHT"; //Value2 is not a typo

            var range_C = sheet.get_Range("C1", "C1");
            range_C.Value2 = "data"; //Value2 is not a typo



            var range_D = sheet.get_Range("D1", "D1");
            range_D.Value2 = "LEFT"; //Value2 is not a typo

            //var range_E = sheet.get_Range("E1", "E1");
            //range_E.Value2 = "mm"; //Value2 is not a typo

            var range_E = sheet.get_Range("E1", "E1");
            range_E.Value2 = "data"; //Value2 is not a typo







            //now the list
            string cellName1, cellName2;
            int counter = 2;
            int idx = 1;

            foreach (var item in list)
            {
                cellName1 = "A" + counter.ToString();
                var range1 = sheet.get_Range(cellName1, cellName1);
                range1.Value2 = item.Item2.ToString();

                cellName2 = "C" + counter.ToString();
                var range2 = sheet.get_Range(cellName2, cellName2);
                range2.Value2 = item.Item3.ToString();



                ++idx;
                ++counter;

                //if (item.Item1 == "400")
                //{
                //   break;
                //}

            }


            string cellName3, cellName4;
            int counter_left = counter;
            int idx_left = 1;

            foreach (var item in list2)
            {
                cellName3 = "A" + counter_left.ToString();
                var range1 = sheet.get_Range(cellName3, cellName3);
                range1.Value2 = item.Item2.ToString();

                cellName4 = "E" + counter_left.ToString();
                var range2 = sheet.get_Range(cellName4, cellName4);
                range2.Value2 = item.Item3.ToString();

                ++idx_left;
                ++counter_left;
            }








            string save_file_path = @"C:\result\excel_test\result_data.xlsx";

            if (File.Exists(save_file_path))
            {

                File.Delete(save_file_path);

            }

            workbook.SaveAs(save_file_path);

            //workbook.Open(save_file_path);

            //excapp.Visible = true;
            //if()
            try
            {
                excapp.Visible = true;
                //workbook.Close();
                excapp.Workbooks.Open(save_file_path);
                workbook = excapp.Workbooks.Open(save_file_path);

            }
            catch
            { }

            //if()
            //{
            //    Application.E
            //}
            ;

            //    excapp.Workbooks.Open(save_file_path);
            //    workbook = excapp.Workbooks.Open(save_file_path);


            //workbook.Close();
            //excapp.Quit();

            //you've probably got the point by now, so a detailed explanation about workbook.SaveAs and workbook.Close is not necessary
            //important: if you did not make excel visible terminating your application will terminate excel as well - I tested it
            //but if you did it - to be honest - I don't know how to close the main excel window - maybee somewhere around excapp.Windows or excapp.ActiveWindow
        }



        bool graph_draw = false;

        public void ListToGraph(List<Tuple<string, double, double>> list, List<Tuple<string, double, double>> list2)
        {
            graph_draw = true;

            chart1.Series[0].Name = "Eyebox";
            chart1.ChartAreas[0].AxisX.Title = "mm";
            chart1.ChartAreas[0].AxisY.Title = "value";

            //chart1.Series[1].Name = "L";
            //chart1.ChartAreas[1].AxisX.Title = "mm";
            //chart1.ChartAreas[1].AxisY.Title = "value";


            //Excel.Workbook XlWorkbook
            /*
            Excel.Workbook workbook = new Excel.Workbook();

            //NsExcel.XLWork

            workbook.AddWorksheet("sheetName");
            var ws = workbook.Worksheet("sheetName");

            int row = 1;
            foreach (object item in list)
            {
                ws.Cell("A" + row.ToString()).Value = item.ToString();
                row++;
            }

            workbook.SaveAs(@"C:\result\excel_test\yourExcel.xlsx");
            */

            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.Series[0].MarkerSize = 80;

            //chart1.Series[1].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            //chart1.Series[1].MarkerSize = 80;

            foreach (var item in list)
            {
                /*
                cellName1 = "A" + counter.ToString();
                var range1 = sheet.get_Range(cellName1, cellName1);
                range1.Value2 = item.Item1.ToString();

                cellName2 = "B" + counter.ToString();
                var range2 = sheet.get_Range(cellName2, cellName2);
                range2.Value2 = item.Item2.ToString();

                ++idx;
                ++counter;
                */

                chart1.Series[0].Points.AddXY(item.Item2, item.Item3);

                //chart1.Series[0].Points.AddXY(1, 1);
                //chart1.Series[0].Points.AddXY(2, 2);
                //chart1.Series[0].Points.AddXY(3, 3);
                //chart1.Series[0].Points.AddXY(4, 4);
            }

            foreach (var item in list2)
            {
                /*
                cellName1 = "A" + counter.ToString();
                var range1 = sheet.get_Range(cellName1, cellName1);
                range1.Value2 = item.Item1.ToString();

                cellName2 = "B" + counter.ToString();
                var range2 = sheet.get_Range(cellName2, cellName2);
                range2.Value2 = item.Item2.ToString();

                ++idx;
                ++counter;
                */
                chart1.Series[0].Points.AddXY(item.Item2, item.Item3);

                //chart1.Series[0].Points.AddXY(1, 1);
                //chart1.Series[0].Points.AddXY(2, 2);
                //chart1.Series[0].Points.AddXY(3, 3);
                //chart1.Series[0].Points.AddXY(4, 4);
            }


            chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.White;
            chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.White;

            chart1.ChartAreas[0].AxisX.Minimum = chart1.Series[0].Points[0].XValue;
            double end_point = list2[list2.Count - 1].Item2;

            chart1.ChartAreas[0].AxisX.Maximum = end_point;

        }




        //***********************************************************************************************************************
        //***********************************************************************************************************************









        private async void pixel_data_save_Click(object sender, EventArgs e)
        {


            var img_gray_task = Task.Run(() =>
            {

                //eyebox_searching();

                //while (continue_signal)
                //{
                //    img_convert();

                //if (continue_signal == false)
                //{
                //    break;
                //}
                //}
            });
            await img_gray_task;







            /*
            
            Mat gray = new Mat();

            const string OutVideoFile = "./out.jpg";
            VideoCapture video_L = new VideoCapture(0);

            //frame_l = Cv2.ImEncode("test", frame_l);
            video_L.Read(frame_l);

            Mat[] mv = new Mat[3];
            Mat mask = new Mat();
            Mat src4 = new Mat();

            Cv2.ImShow("src1", frame_l);
            Cv2.ImWrite("./test1.jpg", frame_l);

            Cv2.CvtColor(frame_l, gray, ColorConversionCodes.RGB2GRAY);
            Cv2.ImWrite("./test1_gray.jpg", gray);


            Cv2.CvtColor(frame_l, src4, ColorConversionCodes.BGR2HSV);
            mv = Cv2.Split(frame_l);
            Cv2.CvtColor(frame_l, src4, ColorConversionCodes.HSV2BGR);
            Cv2.InRange(mv[0], new Scalar(40), new Scalar(100), mask);
            Cv2.BitwiseAnd(frame_l, mask.CvtColor(ColorConversionCodes.GRAY2BGR), frame_l);
            //Cv2.BitwiseAnd(frame_l, mask.CvtColor(ColorConversionCodes.BGR2GRAY), frame_l);
            Cv2.ImShow("result", frame_l);

            Cv2.WaitKey(0);

            Cv2.DestroyAllWindows();

            Cv2.Threshold(frame_l, gray, 130, 255, ThresholdTypes.Binary);
            //Cv2.CvtColor(frame_l, gray, ColorConversionCodes.RGB2GRAY);
            //Cv2.
            //frame_l = Cv2.ImDecode(frame_l, ImreadModes.Grayscale);
            // Cv2.ImWrite("./test1.jpg", frame_l);
            //Cv2.ImWrite("./test1.jpg", gray);
            Mat src = Cv2.ImRead("test1.jpg");

            Cv2.CvtColor(src, gray, ColorConversionCodes.HSV2RGB);
            Cv2.ImWrite("./test2.jpg", gray);

            Cv2.CvtColor(src, gray, ColorConversionCodes.RGB2HSV);
            Cv2.ImWrite("./test2.jpg", gray);

            Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2BGRA);
            Cv2.ImWrite("./test2.jpg", gray);

            Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
            Cv2.ImWrite("./test2.jpg", gray);

            //Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
            //Cv2.ImWrite("./test2.jpg", gray);

            Cv2.ImWrite("./test2.jpg", gray);

            Cv2.ImShow("src", src);
            Cv2.ImShow("gray", gray);
            Cv2.WaitKey(10);
            Cv2.DestroyAllWindows();


            Mat gray2 = new Mat();

            Point2f imgcenter = new Point2f(2048 / 2.0f, 1536 / 2.0f);
            Point2f diff = new Point2f();
            Point2f mean = new Point2f();

            VectorOfVec3f circles = new VectorOfVec3f();

            //Mat gray = new Mat();
            //Mat dst = new Mat();




            //VideoWriter videoWriter = new VideoWriter(OutVideoFile, -1, video_L.Fps, dsize);
            
            using (VideoWriter videoWriter = new VideoWriter(OutVideoFile, -1, video_L.Fps, dsize))
            {
                Console.WriteLine("Converting each movie frames...");
                Mat frame = new Mat();
                while (true)
                {
                    // Read image
                    video_L.Read(frame);
                    if (frame.Empty())
                        break;

                    //Console.CursorLeft = 0;
                    //Console.Write("{0} / {1}", video_L.PosFrames, video_L.FrameCount);

                    // grayscale -> canny -> resize
                    Mat gray = new Mat();
                    Mat canny = new Mat();
                    Mat dst = new Mat();

                    Cv2.CvtColor(frame, gray, ColorConversionCodes.BGR2GRAY);
                    gray1 = gray;

                    Cv2.ImShow("gray", gray);
                    if(Cv2.WaitKey(10) >= 0) break;


                    Cv2.Canny(gray, canny, 100, 180);
                    Cv2.Resize(canny, dst, dsize, 0, 0, InterpolationFlags.Linear);
                    // Write mat to VideoWriter
                    videoWriter.Write(dst);
                }
            }
            
            //VideoCapture video_R = new VideoCapture(1);
            cam_opened = true;
            video_L.Set(CaptureProperty.FrameWidth, 2048);
            video_L.Set(CaptureProperty.FrameHeight, 1536);
            //video_L.Set(CaptureProperty.Format, 32);

            //videoWriter.Write(dst);

            //video_L.Set(CaptureProperty.ConvertRgb, -1);
            //video_L.ConvertRgb = true;
            //video_L.gra
            //video_R.Set(CaptureProperty.FrameWidth, 2048);
            //video_R.Set(CaptureProperty.FrameHeight, 1536);

            Cv2.NamedWindow("src", WindowMode.Normal);
            Cv2.ResizeWindow("src", 640, 480);
            //Cv2.NamedWindow("gray", WindowMode.Normal);
            //Cv2.ResizeWindow("gray", 640, 480);
            img_writing = false;
            show_cam = true;


            for (; ; )
            {
                if (!img_writing)
                {
                    //video_L >> frame_l;

                    video_L.Read(frame_l);

                    Cv2.CvtColor(frame_l, frame_l, ColorConversionCodes.BGR2GRAY);
                    Cv2.ImWrite("./gray_scale_capture.jpg", frame_l);

                    //Mat gray1 = new Mat(OutVideoFile, ImreadModes.Grayscale);

                    //frame_l.ImWrite(OutVideoFile, null);

                    //videoWriter.Write(frame_l);

                    //pictureBox1.Image = BitmapConverter.ToBitmap(frame_l);

                    Cv2.ImShow("src", frame_l);

                    if (Cv2.WaitKey(10) >= 0) break;

                    //frame_l.ImWrite(OutVideoFile);

                    //gray1 = Cv2.ImRead(OutVideoFile, ImreadModes.Grayscale);
                    //video_R.Read(frame_r);
                    //video_L.ConvertRgb = true;

                    //System.IntPtr img_data = frame_l.Data;

                    //byte* img_data1 = frame_l.DataPointer[];



                    for (int row = 0; row < frame_l.Rows; row++)
                    {
                        for (int col = 0; col < frame_l.Cols; col++)
                        {
                            byte r, g, b;

                            // byte *data = (byte *)image.Data.ToPointer();
                            byte* data = (byte*)frame_l.DataPointer;

                            b = data[row * frame_l.Step() + col * frame_l.ElemSize() + 0];
                            g = data[row * frame_l.Step() + col * frame_l.ElemSize() + 1];
                            r = data[row * frame_l.Step() + col * frame_l.ElemSize() + 2];

                            data[row * frame_l.Step() + col * frame_l.ElemSize() + 0] = (byte)(255 - b);
                            data[row * frame_l.Step() + col * frame_l.ElemSize() + 1] = (byte)(255 - g);
                            data[row * frame_l.Step() + col * frame_l.ElemSize() + 2] = (byte)(255 - r);

                            //System.IntPtr b = img_data[row * frame_l.Cols * 3 + col * 3];
                            //System.IntPtr g = img_data[row * frame_l.Cols * 3 + col * 3 + 1];
                            //System.IntPtr r = img_data[row * frame_l.Cols * 3 + col * 3 + 2];
                        }
                    }



                    //------------------------------------------------------------------

                    
                    string body = "...[10만 개의 0~9 숫자로 이뤄진 문자열]...";

                    int loopCount = 10000;
                    byte[] bodyContents = Encoding.UTF8.GetBytes(body);

                    for (int i = 0; i < 100; i++)
                    {
                        MemoryStream ms = new MemoryStream();
                        ms.Write(bodyContents, 0, bodyContents.Length);
                        ms.Flush();
                    }
                    
                    //----------------------------------------------------------------

                    //Cv2.CvtColor(gray, frame_l, ColorConversionCodes.BGRA2GRAY);


                    //frame_l = Cv2.ImRead(frame_l, ImreadModes.Color);
                    //frame_l.i
                    Cv2.Flip(frame_l, frame_l, FlipMode.XY);
                    //Cv2.Flip(frame_r, frame_r, FlipMode.XY);
                    img_updated = true;

                    //Thread.Sleep(500);

                }

                //find_circle_position();
                //Cv2.CvtColor(frame_l, gray, ColorConversionCodes.BGR2GRAY);
                //Thread.Sleep(500);

                //Cv2.MedianBlur(frame_l, frame_l, 5);
                //Cv2.HoughCircles(gray2, HoughMethods.Gradient, 1, 10, 100, 600, 30, 60);




                for (int i = 0; i < circles.Size; i++)
                {

                    //Point center(cvRound(circles[i][0]), cvRound(circles[i][1]));
                    //circle(gray, center, 8, Scalar(0, 255, 0), -1, 8, 0);
                }
                Cv2.ImShow("src", frame_l);
                if (Cv2.WaitKey(10) >= 0 || !show_cam || circles.Size > 5)
                {
                    // t->Abort();
                    // serialprint("L:E");
                    // notfound = false;
                    break;
                }



                Cv2.ImShow("src", frame_l);
                //Cv2.ImShow("gray1", gray1);
                //Cv2.ImShow("gray2", gray2);

                try
                {
                    if (Cv2.WaitKey(10) >= 0 || !show_cam)
                        break;
                }
                catch (Exception ex) { }

                //Cv2.WaitKey(0);
                //Cv2.DestroyAllWindows();






            }

            frame_l.Dispose();
            //gray1.Dispose();

            //frame_r.Dispose();

            video_L.Release();
            //video_R.Release();


            Cv2.DestroyAllWindows();

            */


        }


        private int find_circle_position()
        {
            Mat gray = new Mat();

            //cv::Point2f imgcenter(2048 / 2.0f, 1536 / 2.0f);
            //cv::Point2f diff, mean;
            //std::vector<Vec3f> circles;

            //Point2f 

            Point2f imgcenter = new Point2f(2048 / 2.0f, 1536 / 2.0f);
            Point2f diff = new Point2f();
            Point2f mean = new Point2f();



            //VectorOfVectorPoint2f diff = new VectorOfVectorPoint2f();
            //VectorOfVectorPoint2f mean = new VectorOfVectorPoint2f();

            VectorOfVec3f circles = new VectorOfVec3f();

            Mat shared_lframe_1;

            //Random randomObj = new Random();
            //int randomValue = randomObj.Next();

            bool notfound = true;
            int ncircles = 0;
            double delay = 3000;


            //Cv2.Circle(frame_l, new OpenCvSharp.Point(90, 70), 25, Scalar.Green, -1, LineTypes.AntiAlias);


            //**************Stage Move Command*************************************

            //Cv2.NamedWindow("Searching circles...", WindowMode.Normal);
            //Cv2.ResizeWindow("Searching circles...", 640, 480);

            img_updated = false;
            while (!img_updated) { } // wait the latest frame
            img_writing = true; // lock the shared frames


            Cv2.CvtColor(frame_l, gray, ColorConversionCodes.BGR2GRAY);

            Cv2.ImShow("src", frame_l);
            Cv2.ImShow("gray", gray);
            Cv2.WaitKey(0);
            Cv2.DestroyAllWindows();


            /*        



            for(long i = 0; i < circles.Size; i++)
            {

            }


            Mat[] splitImg = new Mat[3];
            Cv2.Split(shared_lframe, splitImg);

            Vec
            Vector<Mat> mergeImg;
            mergeImg.push_back(splitImg[2]);

            Cv2.Merge(mergeImg, gray);

            splitImg[2].Release();

            img_writing = false;

            Cv2.MedianBlur(gray, gray, 5);


            for (size_t i = 0; i < circles.size(); i++)
            {
                cv::Point center(cvRound(circles[i][0]), cvRound(circles[i][1]));
                circle(gray, center, 8, Scalar(0, 255, 0), -1, 8, 0);
            }
            imshow("Searching circles...", gray);
            if (waitKey(10) >= 0 || !show_cam || circles.size() > 5)
            {
                t->Abort();
                serialprint("L:E");
                notfound = false;
                break;
            }


            */

            return 1;

        }


    }
}
