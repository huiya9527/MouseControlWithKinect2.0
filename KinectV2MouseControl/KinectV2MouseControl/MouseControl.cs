﻿using System;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.Kinect;
using System.Collections;
using System.Collections.Generic;

namespace KinectV2MouseControl
{
    class MouseControl
    {
        private static HandsState beforeHands;
        private static HandsState nowHand;
        private static ArrayList HandsStateList = new ArrayList();

        private const float cursorSmoothing = 0.2f;
        private const int STATE_RECORD_NUM = 5;
        private static int wheeldy = 0;


        private static int screenWidth = (int)SystemParameters.PrimaryScreenWidth;
        private static int screenHeight = (int)SystemParameters.PrimaryScreenHeight;


        private static void MouseLeftDown()
        {
            mouse_event(MouseEventFlag.LeftDown, 0, 0, 0, UIntPtr.Zero);
            Console.WriteLine("mouse left down");
        }
        private static void MouseLeftUp()
        {
            mouse_event(MouseEventFlag.LeftUp, 0, 0, 0, UIntPtr.Zero);
            Console.WriteLine("mouse left up");
        }

        private static void MouseRightDown()
        {
            mouse_event(MouseEventFlag.RightDown, 0, 0, 0, UIntPtr.Zero);
            Console.WriteLine("mouse right down");
        }

        private static void MouseRightUp()
        {
            mouse_event(MouseEventFlag.RightUp, 0, 0, 0, UIntPtr.Zero);
            Console.WriteLine("mouse right up");
        }

        private static void MouseMiddleDown()
        {
            mouse_event(MouseEventFlag.MiddleDown, 0, 0, 0, UIntPtr.Zero);
            Console.WriteLine("mouse middle down");
        }

        private static void MouseMiddleUp()
        {
            mouse_event(MouseEventFlag.MiddleUp, 0, 0, 0, UIntPtr.Zero);
            Console.WriteLine("mouse middle up");
        }

        private static void MouseRoll(int dy)
        {
            mouse_event(MouseEventFlag.Wheel, 0, 0, dy, UIntPtr.Zero);
        }
        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        static extern void mouse_event(MouseEventFlag flags, int dx, int dy, int data, UIntPtr extraInfo);

        /// <summary>
        /// Struct representing a point.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;

            public static implicit operator Point(POINT point)
            {
                return new Point(point.X, point.Y);
            }
        }

        /// <summary>
        /// Retrieves the cursor's position, in screen coordinates.
        /// </summary>
        /// <see>See MSDN documentation for further information.</see>
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        private static Point GetCursorPosition()
        {
            POINT lpPoint;
            GetCursorPos(out lpPoint);
            return lpPoint;
        }

        private static void StateClear()
        {
            if (beforeHands.operation == Operation.left_down)
            {
                MouseLeftUp();
            }
            else if (beforeHands.operation == Operation.middle_down)
            {
                MouseMiddleUp();
            }
            else if (beforeHands.operation == Operation.right_down)
            {
                MouseRightUp();
            }
        }

        private static void move(float mouseSensitivity)
        {
            float x = nowHand.selectHand.X - nowHand.spineBase.X + 0.05f;
            float y = nowHand.spineBase.Y - nowHand.selectHand.Y + 0.51f;
            // get current cursor position
            Point curPos = MouseControl.GetCursorPosition();
            // smoothing for using should be 0 - 0.95f. The way we smooth the cusor is: oldPos + (newPos - oldPos) * smoothValue
            float smoothing = 1 - cursorSmoothing;
            // set cursor position
            MouseControl.SetCursorPos((int)(curPos.X + (x * mouseSensitivity * screenWidth - curPos.X) * smoothing), (int)(curPos.Y + ((y + 0.25f) * mouseSensitivity * screenHeight - curPos.Y) * smoothing));

        }

        private static HandsState findState()
        {
            Dictionary<Operation, int> dic =
            new Dictionary<Operation, int>();

            for (int i = 0; i < STATE_RECORD_NUM; i++)
            {
                HandsState temp = (HandsState)HandsStateList[i];
                if (dic.ContainsKey(temp.operation))
                {
                    dic.Add(temp.operation, dic[temp.operation] + 1);
                }
                else
                {
                    dic.Add(temp.operation, 1);
                }
            }

            int num = 0;
            Operation maxOperation = Operation.no_operation;
            foreach (KeyValuePair<Operation, int> kvp in dic)
            {
                if(kvp.Value >= num)
                {
                    num = kvp.Value;
                    maxOperation = kvp.Key;
                }
            }

            for(int i = 4; i >= 0; i++)
            {
                HandsState temp = (HandsState)HandsStateList[i];
                if (temp.operation == maxOperation)
                {
                    return temp;
                }
            }
            //should not be here;
            return null;
        }

        public static void Mouse_Driver(Body body)
        {
            if (nowHand != null)
            {
                beforeHands = nowHand;
                if (HandsStateList.Count < STATE_RECORD_NUM)
                {
                    nowHand = new HandsState(body);
                }
                else
                {
                    HandsStateList.RemoveAt(0);
                    HandsStateList.Add(new HandsState(body));
                    nowHand = findState();
                }
            }
            else
            {
                beforeHands = new HandsState();
                nowHand = new HandsState(body);
                HandsStateList.Add(nowHand);
            }
            //operation start
            if (nowHand.operation == Operation.no_operation)
            {
                StateClear();
                Console.WriteLine("no operation");
            }
            else if (nowHand.operation == Operation.left_down)
            {
                move(3.5f);
                if (beforeHands.operation == Operation.left_down)
                {
                    return;
                }
                else if (beforeHands.operation == Operation.middle_down)
                {
                    MouseMiddleUp();
                }
                else if (beforeHands.operation == Operation.right_down)
                {  
                    MouseRightUp();
                }
                MouseLeftDown();
            }
            else if (nowHand.operation == Operation.middle_down)
            {
                move(3.5f);
                if (beforeHands.operation == Operation.left_down)
                {
                    MouseLeftUp();
                }
                else if(beforeHands.operation == Operation.middle_down)
                {
                    return;
                }
                else if (beforeHands.operation == Operation.right_down)
                {
                    MouseRightUp();
                }
                MouseMiddleDown();   
            }
            else if (nowHand.operation == Operation.right_down)
            {
                move(3.5f);
                if (beforeHands.operation == Operation.left_down)
                {
                    MouseLeftUp();
                }
                else if (beforeHands.operation == Operation.middle_down)
                {
                    MouseMiddleUp();
                }
                else if(beforeHands.operation == Operation.right_down)
                {
                    return;
                }
                MouseRightDown();
            }
            else if (nowHand.operation == Operation.move)
            {
                StateClear();
                move(3.5f);
            }
            else if(nowHand.operation == Operation.wheel)
            {
                StateClear();
                //wheel
                if(beforeHands.operation != Operation.wheel)
                {
                    wheeldy = beforeHands.primeHandy;
                }
                int dy = nowHand.primeHandy - wheeldy;
                //Console.WriteLine(dy);
                MouseRoll(dy);
            }
        }

        class HandsState
        {
            private const float TOUCH_REGION = 0.15f;
            private const float EMBED_REGION = 0.40f;

            private HandPositionZ LeftHandPosition;
            private HandPositionZ RightHandPosition;
            private HandState LeftHandState;
            private HandState RightHandState;
            private HandPositionZ SelectHandPosition;
            private HandState SelectHandState;
            //for move
            // public CameraSpacePoint selectHand;
            //change whist
            public CameraSpacePoint selectHand;
            public CameraSpacePoint spineBase;
            //for compare
            public Operation operation;
            //for wheel
            public int primeHandy;

            public HandsState(Body body)
            {
                CameraSpacePoint handLeft = body.Joints[JointType.HandLeft].Position;
                CameraSpacePoint handRight = body.Joints[JointType.HandRight].Position;
                CameraSpacePoint wristLeft = body.Joints[JointType.WristLeft].Position;
                CameraSpacePoint wristRight = body.Joints[JointType.WristRight].Position;

                spineBase = body.Joints[JointType.SpineBase].Position;
                primeHandy = (int)(handRight.Y*100);
                //set left hand position
                if (spineBase.Z - handLeft.Z > EMBED_REGION)
                {
                    LeftHandPosition = HandPositionZ.EMBED;
                }
                else if (spineBase.Z - handLeft.Z > TOUCH_REGION)
                {
                    LeftHandPosition = HandPositionZ.TOUCH;
                }
                else
                {
                    LeftHandPosition = HandPositionZ.UNKNOW;
                }
                //set right hand position
                if (spineBase.Z - handRight.Z > EMBED_REGION)
                {
                    RightHandPosition = HandPositionZ.EMBED;
                }
                else if (spineBase.Z - handRight.Z > TOUCH_REGION)
                {
                    RightHandPosition = HandPositionZ.TOUCH;
                }
                else
                {
                    RightHandPosition = HandPositionZ.UNKNOW;
                }
                //set left hand state
                LeftHandState = body.HandLeftState;
                //set right hand state
                RightHandState = body.HandRightState;

                //no hand
                if (LeftHandPosition == HandPositionZ.UNKNOW && RightHandPosition == HandPositionZ.UNKNOW)
                {
                    operation = Operation.no_operation;
                }
                //single hand
                else if (LeftHandPosition == HandPositionZ.UNKNOW || RightHandPosition == HandPositionZ.UNKNOW)
                {
                    //left hand operate
                    if (LeftHandPosition != HandPositionZ.UNKNOW)
                    {
                        selectHand = wristLeft;
                        SelectHandPosition = LeftHandPosition;
                        SelectHandState = LeftHandState;
                    }
                    //right hand operate
                    else
                    {
                        selectHand = wristRight;
                        SelectHandPosition = RightHandPosition;
                        SelectHandState = RightHandState;
                    }
                    //single hand touch region
                    if (SelectHandPosition == HandPositionZ.TOUCH)
                    {
                        if (SelectHandState == HandState.Closed)
                        {
                            operation = Operation.left_down;
                        }
                        else
                        {
                            operation = Operation.move;
                        }
                    }
                    //single hand embed region
                    else
                    {
                        if (SelectHandState == HandState.Closed)
                        {
                            operation = Operation.right_down;
                        }
                        else
                        {
                            operation = Operation.move;
                        }
                    }
                }
                //two hand the prime hand is right
                else
                {
                    //select wrist
                    selectHand = wristRight;
                    //two hand closed will operate wheel
                    if (LeftHandState == HandState.Closed && RightHandState == HandState.Closed)
                    {
                        operation = Operation.wheel;
                    }
                    //one hand closed 
                    else if (LeftHandState == HandState.Closed || RightHandState == HandState.Closed)
                    {
                        operation = Operation.middle_down;
                    }
                    else
                    {
                        operation = Operation.move;
                    }
                }
            }

            public HandsState()
            {
                operation = Operation.no_operation;
            }
        }

        public enum HandPositionZ
        {
            UNKNOW = 0,
            TOUCH = 1,
            EMBED = 2,
        }

        public enum Operation
        {
            no_operation = 0,
            left_down = 1,
            middle_down = 2,
            right_down = 3,
            move = 4,
            wheel = 5,
        }

        [Flags]
        public enum MouseEventFlag : uint
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            XDown = 0x0080,
            XUp = 0x0100,
            Wheel = 0x0800,
            VirtualDesk = 0x4000,
            Absolute = 0x8000
        }
    }
}
