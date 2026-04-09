using libmsstyle;
using msstyleEditor.PropView;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace msstyleEditor.Dialogs
{
    public partial class ClassViewWindow : ToolWindow
    {
        private VisualStyle m_style = null;

        public event TreeViewEventHandler OnSelectionChanged;

        public ClassViewWindow()
        {
            InitializeComponent();
            classView.TreeViewNodeSorter = new TreeRootSorter();
        }

        public void SetVisualStyle(VisualStyle style)
        {
            m_style = style;

            classView.BeginUpdate();
            classView.Nodes.Clear();
            if (m_style != null)
            {
                foreach (var cls in m_style.Classes)
                {
                    var clsNode = new TreeNode(cls.Value.ClassName);
                    clsNode.Tag = cls.Value;
                    if (cls.Value.ClassName == "timingfunction")
                    {
                        foreach (TimingFunction timingFunction in m_style.TimingFunctions)
                        {
                            string name = VisualStyleAnimations.FindTimingFuncName(timingFunction.Header.partID);
                            if(name == null)
                            {
                                name = "Unknown part:" + timingFunction.Header.partID;
                            }

                            var partNode = new TreeNode(name);
                            partNode.Tag = timingFunction;
                            clsNode.Nodes.Add(partNode);
                        }
                    }
                    else if (cls.Value.ClassName == "animations")
                    {
                        // Contains <PartID, partid node>
                        Dictionary<int, TreeNode> nodes = new Dictionary<int, TreeNode>();
                        foreach (Animation animation in m_style.Animations)
                        {
                            var anim = VisualStyleAnimations.FindAnimStates(animation.Header.partID);

                            TreeNode partNode;
                            bool exists = nodes.TryGetValue(animation.Header.partID, out partNode);
                            if(!exists)
                            {
                                // Check if the animations name is known
                                if (anim != null)
                                {
                                    partNode = new TreeNode(anim.AnimationName);
                                }
                                else
                                {
                                    partNode = new TreeNode("Unknown part " + animation.Header.partID);
                                }
                                partNode.Tag = new AnimationTypeDescriptor(animation);
                            }

                            // Find the animation states name
                            string stateName;
                            if(anim != null)
                            {
                                if(!anim.AnimationStateDict.TryGetValue(animation.Header.stateID, out stateName))
                                {
                                    stateName = "Unknown state: " + animation.Header.stateID;
                                }
                            }
                            else
                            {
                                stateName = "Unknown state: " + animation.Header.stateID;
                            }

                            //add the new state if there is another state
                            if (exists)
                            {
                                ((AnimationTypeDescriptor)partNode.Tag).AddState(animation);
                            }

                            // Add the part node if it wasn't added
                            if (!exists)
                            {
                                nodes.Add(animation.Header.partID, partNode);
                                clsNode.Nodes.Add(partNode);
                            }

                        }
                    }
                    else
                    {
                        foreach (var part in cls.Value.Parts)
                        {
                            var partNode = new TreeNode(part.Value.PartName);
                            partNode.Tag = part.Value;

                            foreach (var state in part.Value.States)
                            {
                                var stateNode = new TreeNode(state.Value.StateName);
                                stateNode.Tag = state.Value;
                                partNode.Nodes.Add(stateNode);
                            }

                            clsNode.Nodes.Add(partNode);
                        }
                    }

                    classView.Nodes.Add(clsNode);
                }
                classView.Sort();
            }
            classView.EndUpdate();
        }

        bool m_endReached = false;
        public TreeNode FindNextNode(bool includeSelectedNode, Predicate<TreeNode> predicate)
        {
            var startItem = classView.SelectedNode;
            if (startItem == null || m_endReached)
            {
                m_endReached = false;
                if (classView.Nodes.Count > 0)
                {
                    startItem = classView.Nodes[0];
                }
                else return null;
            }

            var node = TreeViewSearch.FindNextNode(classView, startItem, includeSelectedNode, predicate);
            if (node != null)
            {
                // Because we may visit a class/part node multiple times AND change some properties there, we
                // have to force the `OnSelectionChanged` event to implicitly update the property view.
                bool forceUpdate = classView.SelectedNode == node;
                classView.SelectedNode = node;
                if (forceUpdate)
                {
                    OnTreeItemSelected(this, new TreeViewEventArgs(node));
                }
            }
            else m_endReached = true;

            return node;
        }

        public void ExpandAll()
        {
            classView.ExpandAll();
        }

        public void CollapseAll()
        {
            classView.CollapseAll();
        }

        private void OnTreeItemSelected(object sender, TreeViewEventArgs e)
        {
            if (OnSelectionChanged != null)
            {
                OnSelectionChanged(sender, e);
            }
        }

        public TreeNode AddClassNode(StyleClass cls)
        {
            var clsNode = new TreeNode(cls.ClassName);
            clsNode.Tag = cls;

            // Mirror the regular class tree population so newly added classes
            // immediately show their default parts/states.
            foreach (var part in cls.Parts)
            {
                var partNode = new TreeNode(part.Value.PartName);
                partNode.Tag = part.Value;

                foreach (var state in part.Value.States)
                {
                    var stateNode = new TreeNode(state.Value.StateName);
                    stateNode.Tag = state.Value;
                    partNode.Nodes.Add(stateNode);
                }

                clsNode.Nodes.Add(partNode);
            }

            classView.Nodes.Add(clsNode);
            classView.Sort();
            clsNode.Expand();
            classView.SelectedNode = clsNode;
            return clsNode;
        }

        public TreeNode AddPartNode(TreeNode classNode, StylePart part)
        {
            var partNode = new TreeNode(part.PartName);
            partNode.Tag = part;

            foreach (var state in part.States)
            {
                var stateNode = new TreeNode(state.Value.StateName);
                stateNode.Tag = state.Value;
                partNode.Nodes.Add(stateNode);
            }

            classNode.Nodes.Add(partNode);
            classNode.Expand();
            classView.SelectedNode = partNode;
            return partNode;
        }

        public TreeNode AddStateNode(TreeNode partNode, StyleState state)
        {
            var stateNode = new TreeNode(state.StateName);
            stateNode.Tag = state;
            partNode.Nodes.Add(stateNode);
            partNode.Expand();
            classView.SelectedNode = stateNode;
            return stateNode;
        }

        public TreeNode SelectedNode
        {
            get { return classView.SelectedNode; }
        }
    }
}
