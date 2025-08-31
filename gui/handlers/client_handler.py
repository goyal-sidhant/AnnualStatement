# gui/handlers/client_handler.py
"""Client selection and management operations"""

import os
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, filedialog
from utils.constants import GUI_CONFIG
from utils.helpers import get_state_code


class ClientHandler:
    """Handles client selection and tree operations"""
    
    def __init__(self, app_instance):
        self.app = app_instance
    
    def update_client_tree(self):
        """Update client tree display"""
        if not hasattr(self.app, 'client_tree'):
            return
        
        # Clear existing
        for item in self.app.client_tree.get_children():
            self.app.client_tree.delete(item)
        
        self.app.selected_clients = {}
        
        # Add clients
        for client_key, client_info in self.app.client_data.items():
            missing = ', '.join(client_info['missing_files']) if client_info['missing_files'] else 'None'
            extra = str(len(client_info['extra_files'])) if client_info['extra_files'] else '0'
            
            item = self.app.client_tree.insert('', 'end',
                                        text='☐',
                                        values=(
                                            client_info['client'],
                                            client_info['state'],
                                            f"{client_info.get('status_icon', '')} {client_info['status']}",
                                            client_info['file_count'],
                                            missing,
                                            extra,
                                            '☐ No'
                                        ))
            
            self.app.selected_clients[item] = False
            
            # Initialize folder name setting
            self.app.client_folder_settings[client_key] = False
            
            # Color coding
            if client_info['status'] == 'Complete':
                self.app.client_tree.item(item, tags=('complete',))
            # Ensure text color is correct
            if not self.app.dark_mode.get():
                self.app.client_tree.item(item, tags=('complete',))
                # Force text to be black in light mode
                for col in self.app.client_tree['columns']:
                    self.app.client_tree.set(item, col, self.app.client_tree.set(item, col))
            else:
                self.app.client_tree.item(item, tags=('incomplete',))
        
        # Configure tags
        self.app.client_tree.tag_configure('complete', background=GUI_CONFIG['colors']['complete'])
        self.app.client_tree.tag_configure('incomplete', background=GUI_CONFIG['colors']['incomplete'])
        
        # Select first item if exists
        children = self.app.client_tree.get_children()
        if children:
            self.app.client_tree.selection_set(children[0])
            self.app.client_tree.focus(children[0])
    
    def on_client_click(self, event):
        """Handle client tree click"""
        region = self.app.client_tree.identify_region(event.x, event.y)
        column = self.app.client_tree.identify_column(event.x)
        
        if region == 'tree':
            item = self.app.client_tree.identify_row(event.y)
            if item:
                self.toggle_client_selection(item)
        elif region == 'cell' and column == '#7':  # FolderName column
            item = self.app.client_tree.identify_row(event.y)
            if item:
                self.toggle_folder_name_setting(item)
    
    def on_space_key(self, event):
        """Handle space key for selection"""
        selected = self.app.client_tree.selection()
        if selected:
            self.toggle_client_selection(selected[0])
        return 'break'
    
    def on_up_key(self, event):
        """Handle up arrow key"""
        selected = self.app.client_tree.selection()
        if selected:
            current = selected[0]
            prev_item = self.app.client_tree.prev(current)
            if prev_item:
                self.app.client_tree.selection_set(prev_item)
                self.app.client_tree.focus(prev_item)
                self.app.client_tree.see(prev_item)
        return 'break'
    
    def on_down_key(self, event):
        """Handle down arrow key"""
        selected = self.app.client_tree.selection()
        if selected:
            current = selected[0]
            next_item = self.app.client_tree.next(current)
            if next_item:
                self.app.client_tree.selection_set(next_item)
                self.app.client_tree.focus(next_item)
                self.app.client_tree.see(next_item)
        else:
            # Select first if nothing selected
            children = self.app.client_tree.get_children()
            if children:
                self.app.client_tree.selection_set(children[0])
                self.app.client_tree.focus(children[0])
        return 'break'
    
    def toggle_client_selection(self, item):
        """Toggle client selection"""
        current = self.app.selected_clients.get(item, False)
        self.app.selected_clients[item] = not current
        
        # Update checkbox
        if self.app.selected_clients[item]:
            self.app.client_tree.item(item, text='☑')
        else:
            self.app.client_tree.item(item, text='☐')
        
        # Update status
        selected_count = sum(self.app.selected_clients.values())
        self.app.update_status(f"Selected {selected_count} clients")
    
    def select_all_clients(self):
        """Select all clients"""
        for item in self.app.client_tree.get_children():
            self.app.selected_clients[item] = True
            self.app.client_tree.item(item, text='☑')
        self.app.update_status(f"Selected all {len(self.app.client_data)} clients")
    
    def clear_all_clients(self):
        """Clear all selections"""
        for item in self.app.client_tree.get_children():
            self.app.selected_clients[item] = False
            self.app.client_tree.item(item, text='☐')
        self.app.update_status("Cleared all selections")
    
    def select_complete_clients(self):
        """Select only complete clients"""
        count = 0
        for item in self.app.client_tree.get_children():
            values = self.app.client_tree.item(item)['values']
            if 'Complete' in str(values[2]):
                self.app.selected_clients[item] = True
                self.app.client_tree.item(item, text='☑')
                count += 1
            else:
                self.app.selected_clients[item] = False
                self.app.client_tree.item(item, text='☐')
        self.app.update_status(f"Selected {count} complete clients")
    
    def get_selected_clients(self):
        """Get list of selected clients"""
        selected = []
        
        for item, is_selected in self.app.selected_clients.items():
            if is_selected:
                values = self.app.client_tree.item(item)['values']
                # Convert state name to state code for consistent key
                state_code = get_state_code(values[1])
                client_key = f"{values[0]}-{state_code}"
                selected.append(client_key)
        
        return selected
    
    def toggle_folder_name_setting(self, item):
        """Toggle folder name setting for a client"""
        values = list(self.app.client_tree.item(item)['values'])
        client_key = f"{values[0]}-{values[1]}"
        
        # Toggle the setting
        current = self.app.client_folder_settings.get(client_key, False)
        self.app.client_folder_settings[client_key] = not current
        
        # Update display
        if self.app.client_folder_settings[client_key]:
            values[6] = '☑ Yes'
            # Check name length
            if len(values[0]) > 10:
                response = messagebox.askyesno(
                    "Long Client Name Warning",
                    f"Client name '{values[0]}' has {len(values[0])} characters.\n\n"
                    "This will create long folder names. Continue?"
                )
                if not response:
                    self.app.client_folder_settings[client_key] = False
                    values[6] = '☐ No'
        else:
            values[6] = '☐ No'
            # If unchecking and global is ON, turn off global
            if self.app.include_client_name_in_folders.get():
                self.app.include_client_name_in_folders.set(False)
                messagebox.showinfo("Global Setting Disabled", 
                                "Global folder name setting has been turned off\n"
                                "since you unchecked an individual client.")
        
        self.app.client_tree.item(item, values=values)
    
    def update_global_folder_setting(self):
        """Update all client folder settings when global checkbox changes"""
        if self.app.include_client_name_in_folders.get():
            # Global is ON - update all individual settings to Yes
            for item in self.app.client_tree.get_children():
                values = list(self.app.client_tree.item(item)['values'])
                client_key = f"{values[0]}-{values[1]}"
                self.app.client_folder_settings[client_key] = True
                values[6] = '☑ Yes'
                self.app.client_tree.item(item, values=values)
        # Don't change individual settings when global is turned OFF
        self.app.save_cache()
    
    def export_client_list(self):
        """Export selected clients to text file"""
        selected = self.get_selected_clients()
        if not selected:
            messagebox.showwarning("No Selection", "Please select clients to export")
            return
        
        # Ask where to save
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"selected_clients_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(f"Selected Clients List\n")
                    f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"Total: {len(selected)} clients\n")
                    f.write("=" * 50 + "\n\n")
                    
                    for client_key in selected:
                        client_info = self.app.client_data[client_key]
                        f.write(f"{client_info['client']} - {client_info['state']}\n")
                
                messagebox.showinfo("Export Complete", 
                                f"Exported {len(selected)} clients to:\n{os.path.basename(filename)}")
                
                # Open the file
                os.startfile(filename)
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export: {str(e)}")