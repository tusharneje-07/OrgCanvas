"""
Tree Builder Module
Converts flat employee list into hierarchical tree structure.

The algorithm works as follows:
1. Create a lookup dictionary mapping employee_id to employee data
2. Add a 'children' array to each employee
3. Iterate through all employees:
   - If employee has no manager_id, they are a root node
   - If employee has a manager_id, add them to their manager's children array
4. Return the root node(s) which now contain the full tree

This approach has O(n) time complexity where n is the number of employees.
"""

from typing import Dict, List, Any, Optional
import copy
import logging

logger = logging.getLogger(__name__)


class TreeBuilderError(Exception):
    """Custom exception for tree building errors."""
    pass


class OrgTreeNode:
    """
    Represents a node in the organizational tree.
    
    Attributes:
        id (str): Unique employee identifier
        name (str): Employee name
        title (str): Job title
        department (str): Department name
        manager_id (Optional[str]): ID of the manager
        avatar_url (Optional[str]): URL to avatar image
        color (str): Node color (hex)
        children (List[OrgTreeNode]): Child nodes (direct reports)
    """
    
    def __init__(self, data: Dict[str, Any]):
        """
        Initialize a tree node from employee data.
        
        Args:
            data: Employee dictionary containing node data
        """
        self.id = str(data.get('id', ''))
        self.name = data.get('name', '')
        self.title = data.get('title', '')
        self.department = data.get('department', '')
        self.manager_id = data.get('manager_id')
        self.avatar_url = data.get('avatar_url')
        self.color = data.get('color', '#757575')
        self.whatsapp = data.get('whatsapp')
        self.children: List['OrgTreeNode'] = []
        
        # UI state properties (for frontend)
        self.expanded = True
        self.x = 0
        self.y = 0
    
    def to_dict(self) -> Dict[str, Any]:
        """
        Convert node and all children to dictionary format.
        
        Returns:
            Dictionary representation of the node tree
        """
        return {
            'id': self.id,
            'name': self.name,
            'title': self.title,
            'department': self.department,
            'manager_id': self.manager_id,
            'avatar_url': self.avatar_url,
            'color': self.color,
            'whatsapp': self.whatsapp,
            'expanded': self.expanded,
            'children': [child.to_dict() for child in self.children]
        }
    
    def add_child(self, child: 'OrgTreeNode') -> None:
        """
        Add a child node to this node's children.
        
        Args:
            child: Child node to add
        """
        self.children.append(child)
    
    def remove_child(self, child_id: str) -> Optional['OrgTreeNode']:
        """
        Remove a child node by ID.
        
        Args:
            child_id: ID of the child to remove
            
        Returns:
            Removed child node or None if not found
        """
        for i, child in enumerate(self.children):
            if child.id == child_id:
                return self.children.pop(i)
        return None
    
    def find_node(self, node_id: str) -> Optional['OrgTreeNode']:
        """
        Find a node by ID in this subtree.
        
        Args:
            node_id: ID of the node to find
            
        Returns:
            Found node or None
        """
        if self.id == node_id:
            return self
        
        for child in self.children:
            found = child.find_node(node_id)
            if found:
                return found
        return None
    
    def get_all_descendants(self) -> List['OrgTreeNode']:
        """
        Get all descendant nodes (children, grandchildren, etc.).
        
        Returns:
            Flat list of all descendant nodes
        """
        descendants = []
        for child in self.children:
            descendants.append(child)
            descendants.extend(child.get_all_descendants())
        return descendants
    
    def get_depth(self) -> int:
        """
        Calculate the depth of this subtree.
        
        Returns:
            Maximum depth from this node to leaf nodes
        """
        if not self.children:
            return 1
        return 1 + max(child.get_depth() for child in self.children)
    
    def count_nodes(self) -> int:
        """
        Count total nodes in this subtree.
        
        Returns:
            Total number of nodes including this node
        """
        count = 1
        for child in self.children:
            count += child.count_nodes()
        return count


class TreeBuilder:
    """
    Builds organizational tree from flat employee data.
    
    The tree builder converts a flat list of employees with parent references
    into a hierarchical tree structure that can be easily rendered as an org chart.
    
    Attributes:
        employees (List[Dict]): Raw employee data
        nodes (Dict[str, OrgTreeNode]): Lookup dictionary of all nodes
        roots (List[OrgTreeNode]): Root nodes of the tree
    """
    
    def __init__(self, employees: List[Dict[str, Any]]):
        """
        Initialize the tree builder.
        
        Args:
            employees: List of employee dictionaries from Excel parser
        """
        self.employees = employees
        self.nodes: Dict[str, OrgTreeNode] = {}
        self.roots: List[OrgTreeNode] = []
    
    def build(self) -> List[OrgTreeNode]:
        """
        Build the organizational tree from flat employee data.
        
        Algorithm:
        1. First pass: Create all nodes and add to lookup dictionary
        2. Second pass: Link children to parents
        3. Identify root nodes (those with no manager or invalid manager)
        
        Returns:
            List of root nodes (typically just one for org charts)
            
        Raises:
            TreeBuilderError: If tree cannot be built
        """
        if not self.employees:
            logger.warning("No employees provided, returning empty tree")
            return []
        
        # First pass: Create all nodes
        # O(n) where n = number of employees
        for employee in self.employees:
            node = OrgTreeNode(employee)
            self.nodes[node.id] = node
        
        # Second pass: Link children to parents
        # O(n) - each employee is processed once
        orphans = []
        for node_id, node in self.nodes.items():
            manager_id = node.manager_id
            
            if manager_id is None or manager_id == '':
                # This is a root node
                self.roots.append(node)
            elif manager_id in self.nodes:
                # Link to parent
                self.nodes[manager_id].add_child(node)
            else:
                # Invalid manager reference - treat as orphan
                logger.warning(
                    f"Employee {node_id} ({node.name}) has invalid manager_id: {manager_id}"
                )
                orphans.append(node)
        
        # Handle orphans - add them as root nodes
        for orphan in orphans:
            self.roots.append(orphan)
        
        # Validate tree structure
        if not self.roots:
            raise TreeBuilderError(
                "No root nodes found. Ensure at least one employee has no manager_id."
            )
        
        # Check for circular references
        self._detect_cycles()
        
        logger.info(
            f"Built org tree with {len(self.roots)} root(s) and "
            f"{len(self.nodes)} total nodes"
        )
        
        return self.roots
    
    def _detect_cycles(self) -> None:
        """
        Detect circular references in the tree.
        
        Raises:
            TreeBuilderError: If a cycle is detected
        """
        visited = set()
        rec_stack = set()
        
        def dfs(node: OrgTreeNode) -> bool:
            visited.add(node.id)
            rec_stack.add(node.id)
            
            for child in node.children:
                if child.id not in visited:
                    if dfs(child):
                        return True
                elif child.id in rec_stack:
                    return True
            
            rec_stack.remove(node.id)
            return False
        
        for root in self.roots:
            if root.id not in visited:
                if dfs(root):
                    raise TreeBuilderError(
                        "Circular reference detected in organizational hierarchy. "
                        "Please check manager assignments."
                    )
    
    def to_dict(self) -> Dict[str, Any]:
        """
        Convert the entire tree to a dictionary.
        
        Returns:
            Dictionary with roots array and metadata
        """
        return {
            'roots': [root.to_dict() for root in self.roots],
            'total_employees': len(self.nodes),
            'max_depth': max((root.get_depth() for root in self.roots), default=0)
        }
    
    def to_flat_list(self) -> List[Dict[str, Any]]:
        """
        Convert tree back to flat list format for saving.
        
        Returns:
            Flat list of employee dictionaries
        """
        flat_list = []
        
        def traverse(node: OrgTreeNode):
            flat_list.append({
                'id': node.id,
                'name': node.name,
                'title': node.title,
                'department': node.department,
                'manager_id': node.manager_id,
                'avatar_url': node.avatar_url,
                'color': node.color
            })
            for child in node.children:
                traverse(child)
        
        for root in self.roots:
            traverse(root)
        
        return flat_list
    
    def find_node(self, node_id: str) -> Optional[OrgTreeNode]:
        """
        Find a node by ID anywhere in the tree.
        
        Args:
            node_id: ID of the node to find
            
        Returns:
            Found node or None
        """
        return self.nodes.get(node_id)
    
    def add_node(self, employee_data: Dict[str, Any]) -> OrgTreeNode:
        """
        Add a new node to the tree.
        
        Args:
            employee_data: Dictionary containing employee data
            
        Returns:
            The newly created node
            
        Raises:
            TreeBuilderError: If node cannot be added
        """
        node = OrgTreeNode(employee_data)
        
        if node.id in self.nodes:
            raise TreeBuilderError(f"Employee with ID {node.id} already exists")
        
        self.nodes[node.id] = node
        
        if node.manager_id and node.manager_id in self.nodes:
            self.nodes[node.manager_id].add_child(node)
        else:
            self.roots.append(node)
        
        return node
    
    def update_node(self, node_id: str, updates: Dict[str, Any]) -> Optional[OrgTreeNode]:
        """
        Update an existing node's data.
        
        Args:
            node_id: ID of the node to update
            updates: Dictionary of fields to update
            
        Returns:
            Updated node or None if not found
        """
        node = self.find_node(node_id)
        if not node:
            return None
        
        # Handle manager change (reparenting)
        if 'manager_id' in updates and updates['manager_id'] != node.manager_id:
            self._reparent_node(node, updates['manager_id'])
        
        # Update other fields
        for key, value in updates.items():
            if key != 'children' and hasattr(node, key):
                setattr(node, key, value)
        
        return node
    
    def _reparent_node(self, node: OrgTreeNode, new_manager_id: Optional[str]) -> None:
        """
        Change a node's parent (manager).
        
        Args:
            node: Node to reparent
            new_manager_id: ID of the new manager (None for root)
        """
        # Remove from current parent
        if node.manager_id and node.manager_id in self.nodes:
            self.nodes[node.manager_id].remove_child(node.id)
        elif node in self.roots:
            self.roots.remove(node)
        
        # Update manager_id
        node.manager_id = new_manager_id
        
        # Add to new parent
        if new_manager_id and new_manager_id in self.nodes:
            self.nodes[new_manager_id].add_child(node)
        else:
            self.roots.append(node)
    
    def delete_node(self, node_id: str, reassign_to: Optional[str] = None) -> bool:
        """
        Delete a node from the tree.
        
        Args:
            node_id: ID of the node to delete
            reassign_to: ID of node to reassign children to (optional)
            
        Returns:
            True if deletion successful
            
        Raises:
            TreeBuilderError: If deletion cannot be performed
        """
        node = self.find_node(node_id)
        if not node:
            raise TreeBuilderError(f"Node with ID {node_id} not found")
        
        # Handle children
        children = list(node.children)  # Copy to avoid modification during iteration
        
        if reassign_to and reassign_to in self.nodes:
            # Reassign children to specified node
            new_parent = self.nodes[reassign_to]
            for child in children:
                child.manager_id = reassign_to
                new_parent.add_child(child)
                node.remove_child(child.id)
        elif node.manager_id and node.manager_id in self.nodes:
            # Reassign children to deleted node's parent
            parent = self.nodes[node.manager_id]
            for child in children:
                child.manager_id = node.manager_id
                parent.add_child(child)
                node.remove_child(child.id)
        else:
            # Children become root nodes
            for child in children:
                child.manager_id = None
                self.roots.append(child)
                node.remove_child(child.id)
        
        # Remove node from parent
        if node.manager_id and node.manager_id in self.nodes:
            self.nodes[node.manager_id].remove_child(node_id)
        elif node in self.roots:
            self.roots.remove(node)
        
        # Remove from nodes dictionary
        del self.nodes[node_id]
        
        return True
    
    def generate_new_id(self) -> str:
        """
        Generate a new unique employee ID.
        
        Returns:
            New unique ID string
        """
        # Find the highest numeric ID and increment
        max_id = 0
        for node_id in self.nodes.keys():
            try:
                num_id = int(node_id)
                max_id = max(max_id, num_id)
            except ValueError:
                continue
        
        return str(max_id + 1)


def build_tree_from_data(employees: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Convenience function to build tree and return dictionary.
    
    Args:
        employees: List of employee dictionaries
        
    Returns:
        Dictionary representation of the organizational tree
    """
    builder = TreeBuilder(employees)
    builder.build()
    return builder.to_dict()
