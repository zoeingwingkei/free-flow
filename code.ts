interface PluginState {
  color: string; // 6-char hex color code without the leading #
  opacity: number; // 0-100
  startMargin: number; // 0-10
  endMargin: number; // 0-10
  weight: number; // 1-10
  radius: number; // 0-100
  dash: number; // 0-10
  dashGap: number; // 0-10
  startCap: StrokeCap;
  endCap: StrokeCap;
  lineStyle: string;
}

interface ArrowStyle {
  color: string;
  opacity: number;
  weight: number;
  radius: number;
  dash: number;
  dashGap: number;
}

interface ArrowName {
  startName: string;
  endName: string;
}

interface ArrowGeometry {
  startMargin: number;
  endMargin: number;
  startCap: StrokeCap;
  endCap: StrokeCap;
  lineStyle: string;
}

interface ArrowPosition {
  startRect: Rect;
  endRect: Rect;
}

interface ArrowText {
  text: string;
}

interface ArrowAnchors {
  startX: number;
  startY: number;
  endX: number;
  endY: number;
}

type ArrowDirection = 'vertical' | 'horizontal' | 'invalid';

interface Arrow {
  id: string;
  startNodeID: string;
  endNodeID: string;
  textNodeID?: string;
  position: ArrowPosition;
  geometry: ArrowGeometry;
  style: ArrowStyle;
  name: ArrowName;
  text?: ArrowText;
  direction: ArrowDirection;
}

/** UTILITY MODULES */
const ColorUtils = {
  hexToFigmaRGB(hex: string): RGB {
    let c = hex;
    if (c.length === 3) c = c.split('').map(x => x + x).join('');
    const num = parseInt(c, 16);
    return {
      r: ((num >> 16) & 255) / 255,
      g: ((num >> 8) & 255) / 255,
      b: (num & 255) / 255
    };
  },

  numberToFigmaOpacity(num: number): number {
    return num / 100;
  },

  getTextColor(hex: string): RGB {

    // Convert hex to RGB
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);

    // Calculate relative luminance using WCAG formula
    // https://www.w3.org/WAI/WCAG21/Understanding/contrast-minimum.html
    const getRGB = (color: number) => {
      const c = color / 255;
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    };

    const rLum = getRGB(r);
    const gLum = getRGB(g);
    const bLum = getRGB(b);

    // Calculate luminance
    const luminance = 0.2126 * rLum + 0.7152 * gLum + 0.0722 * bLum;

    // Calculate contrast ratios for black and white text
    const blackContrast = (luminance + 0.05) / (0 + 0.05);
    const whiteContrast = (1 + 0.05) / (luminance + 0.05);

    // Return the color with higher contrast
    return blackContrast > whiteContrast ? ColorUtils.hexToFigmaRGB("000000") : ColorUtils.hexToFigmaRGB("FFFFFF");
  }
};

const DebounceUtils = {
  debounce<T extends (...args: any[]) => any>(func: T, delay: number): (...args: Parameters<T>) => void {
    let timeoutId: number | null = null;
    return (...args: Parameters<T>) => {
      if (timeoutId !== null) {
        clearTimeout(timeoutId);
      }
      timeoutId = setTimeout(() => func.apply(this, args), delay);
    };
  },
};

const NodeUtils = {
  getAbsoluteBounds(node: SceneNode): Rect {
    return node.absoluteBoundingBox || { x: node.x, y: node.y, width: node.width, height: node.height };
  },

  getAncestors(node: BaseNode): (SectionNode | FrameNode)[] {
    const ancestors: (SectionNode | FrameNode)[] = [];
    let parent = node.parent;
    while (parent) {
      if (parent.type === "SECTION" || parent.type === "FRAME") {
        ancestors.push(parent as SectionNode | FrameNode);
      }
      parent = parent.parent;
    }
    return ancestors;
  },

  findCommonParent(nodeA: SceneNode, nodeB: SceneNode): SectionNode | FrameNode | null {
    const aAncestors = this.getAncestors(nodeA);
    const bAncestors = this.getAncestors(nodeB);

    // Find the first common ancestor that is a section or frame
    for (const ancestor of aAncestors) {
      if (bAncestors.some(b => b.id === ancestor.id)) {
        return ancestor;
      }
    }

    return null;
  },

  appendToContext(arrow: SceneNode, startNode: SceneNode, endNode: SceneNode): void {
    const commonParent = this.findCommonParent(startNode, endNode);

    if (commonParent) {
      // Adjust arrow coordinates relative to the common parent
      const parentBounds = this.getAbsoluteBounds(commonParent as SceneNode);
      arrow.x -= parentBounds.x;
      arrow.y -= parentBounds.y;
      commonParent.appendChild(arrow);
    } else {
      // No common parent found, append to current page
      figma.currentPage.appendChild(arrow);
    }
  },

  async getNodeById(nodeId: string): Promise<BaseNode | null> {
    try {
      return await figma.getNodeByIdAsync(nodeId);
    } catch (error) {
      console.error(`Node with ID ${nodeId} not found:`, error);
      return null;
    }
  },

  async getNodesByIds(nodeIds: string[]): Promise<(BaseNode | null)[]> {
    return Promise.all(nodeIds.map(id => this.getNodeById(id)));
  },

  isNodeSelected(node: SceneNode | null): boolean {
    if (!node) {
      return false;
    }
    return figma.currentPage.selection.some(selectedNode => selectedNode.id === node.id);
  },

  getAllDescendantIds(node: BaseNode): string[] {
    const descendantIds: string[] = [];

    const traverse = (currentNode: BaseNode) => {
      // Add current node's ID
      descendantIds.push(currentNode.id);

      // Check if node has children and traverse them
      if ('children' in currentNode && currentNode.children) {
        for (const child of currentNode.children) {
          traverse(child);
        }
      }
    };

    traverse(node);
    return descendantIds;
  },
}

const GeometryCalculator = {
  getRelativePosition(rect1: Rect, rect2: Rect): string {
    let posResult = "UNKNOWN";

    const margin = 16; // Margin to avoid overlap

    const topY1 = rect1.y - margin;
    const bottomY1 = rect1.y + rect1.height + margin;
    const leftX1 = rect1.x - margin;
    const rightX1 = rect1.x + rect1.width + margin;
    const topY2 = rect2.y;
    const bottomY2 = rect2.y + rect2.height;
    const leftX2 = rect2.x;
    const rightX2 = rect2.x + rect2.width;

    // determine the Y relative position
    if (bottomY2 < topY1) {
      posResult = "T";
    } else if (topY2 > bottomY1) {
      posResult = "B";
    } else {
      posResult = "M";
    }

    // determine the X relative position
    if (rightX2 < leftX1) {
      posResult += "L";
    } else if (leftX2 > rightX1) {
      posResult += "R";
    } else {
      posResult += "M";
    }

    return posResult;
  },

  getDirection(rect1: Rect, rect2: Rect): ArrowDirection {
    const position = this.getRelativePosition(rect1, rect2);
    let direction: ArrowDirection = 'invalid';
    if (position.includes('L') || position.includes('R')) {
      direction = 'horizontal';
    } else if (position.includes('T') || position.includes('B')) {
      direction = 'vertical';
    }
    // console.log('[Free Flow] get Direction: ', direction);
    return direction;
  },

  getValidDirection(direction: ArrowDirection, rect1: Rect, rect2: Rect): ArrowDirection {
    const position = this.getRelativePosition(rect1, rect2);
    let newDirection: ArrowDirection = 'invalid';
    switch (direction) {
      case 'horizontal':
        if (position.includes('L') || position.includes('R')) {
          newDirection = direction;
        } else {
          newDirection = this.getDirection(rect1, rect2);
        }
        break;
      case 'vertical':
        if (position.includes('T') || position.includes('B')) {
          newDirection = direction;
        } else {
          newDirection = this.getDirection(rect1, rect2);
        }
        break;
      default:
        newDirection = this.getDirection(rect1, rect2);
    }
    return newDirection;
  },

  getAnchorPoints(startRect: Rect, endRect: Rect, startMargin: number, endMargin: number, direction: ArrowDirection): ArrowAnchors {
    const position = this.getRelativePosition(startRect, endRect);
    let startX: number, startY: number, endX: number, endY: number;

    switch (direction) {
      case 'horizontal':
        if (position.includes('L')) {
          // End object is to the left
          startX = startRect.x - startMargin;
          startY = startRect.y + startRect.height / 2;
          endX = endRect.x + endRect.width + endMargin;
          endY = endRect.y + endRect.height / 2;
        } else {
          // End object is to the right
          startX = startRect.x + startRect.width + startMargin;
          startY = startRect.y + startRect.height / 2;
          endX = endRect.x - endMargin;
          endY = endRect.y + endRect.height / 2;
        }
        break;
      case 'vertical':
        if (position.includes('T')) {
          // End object is above
          startX = startRect.x + startRect.width / 2;
          startY = startRect.y - startMargin;
          endX = endRect.x + endRect.width / 2;
          endY = endRect.y + endRect.height + endMargin;
        } else {
          // End object is below
          startX = startRect.x + startRect.width / 2;
          startY = startRect.y + startRect.height + startMargin;
          endX = endRect.x + endRect.width / 2;
          endY = endRect.y - endMargin;
        }
        break;
      default:
        throw new Error(`Invalid position for arrow creation: ${position}`);
    }

    return { startX, startY, endX, endY };
  },

  getArrowVertices(startRect: Rect, endRect: Rect, geometry: ArrowGeometry, direction: ArrowDirection): VectorVertex[] {
    const { startMargin, endMargin, startCap, endCap, lineStyle } = geometry;

    const { startX, startY, endX, endY } = this.getAnchorPoints(startRect, endRect, startMargin, endMargin, direction);

    let vertices: VectorVertex[] = [];

    if (lineStyle === 'CURVE') {
      vertices = [
        {
          x: startX,
          y: startY,
          strokeCap: startCap,
        },
        {
          x: endX,
          y: endY,
          strokeCap: endCap,
        }
      ]
    } else if (lineStyle === 'GRID') {
      const position = this.getRelativePosition(startRect, endRect);
      switch (direction) {
        case 'horizontal':
          // Horizontal-first L-shaped arrow
          vertices = [
            {
              x: startX,
              y: startY,
              strokeCap: startCap,
              strokeJoin: 'ROUND'
            },
            {
              x: startX / 2 + endX / 2,
              y: startY,
              strokeCap: 'NONE',
              strokeJoin: 'ROUND'
            },
            {
              x: startX / 2 + endX / 2,
              y: endY,
              strokeCap: 'NONE',
              strokeJoin: 'ROUND'
            },
            {
              x: endX,
              y: endY,
              strokeCap: endCap,
              strokeJoin: 'ROUND'
            }
          ];
          break;

        case 'vertical':
          // Vertical-first L-shaped arrow
          vertices = [
            {
              x: startX,
              y: startY,
              strokeCap: startCap,
              strokeJoin: 'ROUND'
            },
            {
              x: startX,
              y: (startY + endY) / 2,
              strokeCap: 'NONE',
              strokeJoin: 'ROUND'
            },
            {
              x: endX,
              y: (startY + endY) / 2,
              strokeCap: 'NONE',
              strokeJoin: 'ROUND'
            },
            {
              x: endX,
              y: endY,
              strokeCap: endCap,
              strokeJoin: 'ROUND'
            }
          ];
          break;

        default:
          throw new Error(`Cannot create arrow for position: ${position}`);
      }
    } else {
      throw new Error(`Invalid line style: ${lineStyle}`);
    }

    return vertices;
  },

  getArrowSegments(vertexCount: number, startRect: Rect, endRect: Rect, geometry: ArrowGeometry, direction: ArrowDirection): VectorSegment[] {
    const { lineStyle, startMargin, endMargin } = geometry;

    let segments: VectorSegment[] = [];

    if (lineStyle === 'CURVE') {
      const { startX, startY, endX, endY } = this.getAnchorPoints(startRect, endRect, startMargin, endMargin, direction);
      const position = this.getRelativePosition(startRect, endRect);

      switch (direction) {
        case 'horizontal':
          segments = [
            {
              start: 0,
              end: 1,
              tangentStart: { x: (endX - startX) / 2, y: 0 },
              tangentEnd: { x: (startX - endX) / 2, y: 0 }
            }
          ]
          break;

        case 'vertical':
          segments = [
            {
              start: 0,
              end: 1,
              tangentStart: { x: 0, y: (endY - startY) / 2 },
              tangentEnd: { x: 0, y: (startY - endY) / 2 }
            }
          ]
          break;

        default:
          throw new Error(`Invalid position for arrow creation: ${position}`);
      }

    } else if (lineStyle === 'GRID') {
      for (let i = 0; i < vertexCount - 1; i++) {
        segments.push({
          start: i,
          end: i + 1,
          tangentStart: { x: 0, y: 0 },
          tangentEnd: { x: 0, y: 0 }
        });
      }
    } else {
      throw new Error(`Invalid line style: ${lineStyle}`);
    }

    return segments;
  },

}

/** STATE MANAGER */
class State {
  private document: DocumentNode;
  private state: PluginState;
  private static readonly DEFAULTS: PluginState = {
    color: '000000',
    opacity: 100,
    startMargin: 0,
    endMargin: 0,
    weight: 2,
    radius: 12,
    dash: 0,
    dashGap: 0,
    startCap: 'ROUND' as StrokeCap,
    endCap: 'ARROW_LINES' as StrokeCap,
    lineStyle: 'GRID' as string
  } as const;
  private static readonly STORAGE_KEYS = {
    color: 'savedColor',
    opacity: 'savedOpacity',
    startMargin: 'savedStartMargin',
    endMargin: 'savedEndMargin',
    weight: 'savedWeight',
    radius: 'savedRadius',
    dash: 'savedDash',
    dashGap: 'savedDashGap',
    startCap: 'savedStartCap',
    endCap: 'savedEndCap',
    lineStyle: 'savedLineStyle'
  } as const;

  constructor() {
    this.document = figma.root;
    this.state = this.loadState();
  }

  loadState(): PluginState {
    const loadedState: PluginState = {
      color: this.document.getPluginData(State.STORAGE_KEYS.color) || State.DEFAULTS.color,
      opacity: Number(this.document.getPluginData(State.STORAGE_KEYS.opacity)) || State.DEFAULTS.opacity,
      startMargin: Number(this.document.getPluginData(State.STORAGE_KEYS.startMargin)) || State.DEFAULTS.startMargin,
      endMargin: Number(this.document.getPluginData(State.STORAGE_KEYS.endMargin)) || State.DEFAULTS.endMargin,
      weight: Number(this.document.getPluginData(State.STORAGE_KEYS.weight)) || State.DEFAULTS.weight,
      radius: Number(this.document.getPluginData(State.STORAGE_KEYS.radius)) || State.DEFAULTS.radius,
      dash: Number(this.document.getPluginData(State.STORAGE_KEYS.dash)) || State.DEFAULTS.dash,
      dashGap: Number(this.document.getPluginData(State.STORAGE_KEYS.dashGap)) || State.DEFAULTS.dashGap,
      startCap: this.document.getPluginData(State.STORAGE_KEYS.startCap) as StrokeCap || State.DEFAULTS.startCap,
      endCap: this.document.getPluginData(State.STORAGE_KEYS.endCap) as StrokeCap || State.DEFAULTS.endCap,
      lineStyle: this.document.getPluginData(State.STORAGE_KEYS.lineStyle) as string || State.DEFAULTS.lineStyle
    }

    return loadedState;
  }

  updateState(newState: Partial<PluginState>) {
    this.state = { ...this.state, ...newState } as PluginState;

    // Update Plugin Data
    Object.keys(newState).forEach(key => {
      const value = newState[key as keyof PluginState];
      const storageKey = State.STORAGE_KEYS[key as keyof PluginState];
      if (value !== undefined && storageKey !== undefined) {
        this.document.setPluginData(storageKey, value.toString());
      }
    })

  }

  getStyle(): ArrowStyle {
    return {
      color: this.state.color,
      opacity: this.state.opacity,
      weight: this.state.weight,
      radius: this.state.radius,
      dash: this.state.dash,
      dashGap: this.state.dashGap,
    }
  }

  getGeometry(): ArrowGeometry {
    return {
      startMargin: this.state.startMargin,
      endMargin: this.state.endMargin,
      startCap: this.state.startCap,
      endCap: this.state.endCap,
      lineStyle: this.state.lineStyle
    }
  }

  getState(): PluginState {
    return this.state;
  }
}

/** SELECTION TRACKER */
class Selection {
  private firstSelected: SceneNode | null = null;
  private secondSelected: SceneNode | null = null;
  private arrowSelected: boolean = false;
  private currentArrow: VectorNode | null = null;
  private autoDraw: boolean = false;

  constructor() {
  }

  // Get the start node and end node for the arrow
  getStartAndEndNodes(): { startNode: SceneNode, endNode: SceneNode } {
    // Check if exactly 2 objects are selected
    if (figma.currentPage.selection.length !== 2) {
      figma.notify('Please select exactly 2 objects to create an arrow.');
      throw new Error('Please select exactly 2 objects to create an arrow.');
    }

    // If there is first & second selected, return them
    if (this.firstSelected && this.secondSelected) {
      return {
        startNode: this.firstSelected,
        endNode: this.secondSelected
      }
    }
    // Else, return the first 2 random nodes in the selection
    return {
      startNode: figma.currentPage.selection[0],
      endNode: figma.currentPage.selection[1]
    }
  }

  isArrowSelected(): boolean {
    return this.arrowSelected;
  }

  setSelectedArrow(arrow: VectorNode) {
    this.currentArrow = arrow;
    this.arrowSelected = true;
  }

  setFirstSelected(node: SceneNode) {
    this.firstSelected = node;
  }

  setSecondSelected(node: SceneNode) {
    this.secondSelected = node;
  }

  setFirstAndSecondSelected(selection: readonly SceneNode[]) {
    // If first selected node is selected, set the second selected node
    if (NodeUtils.isNodeSelected(this.firstSelected)) {
      const secondSelected = selection.find(node => node.id !== this.firstSelected?.id);
      if (secondSelected) {
        this.setSecondSelected(secondSelected);
        return;
      }
    }

    // Else, set the first and second selected nodes randomly
    this.setFirstSelected(selection[0]);
    this.setSecondSelected(selection[1]);
  }

  clearSelectedArrow() {
    this.currentArrow = null;
    this.arrowSelected = false;
  }

  getSelectedArrow() {
    return this.currentArrow;
  }

  setAutoDraw(autoDraw: boolean) {
    this.autoDraw = autoDraw;
  }

  getAutoDraw() {
    return this.autoDraw;
  }
}

/** ARROW MANAGER */
class ArrowManager {
  arrows: Map<string, Arrow>;
  private updatingArrows: Set<string> = new Set(); // Track arrows being updated
  private removingArrows: Set<string> = new Set(); // Track arrows being removed

  constructor() {
    this.arrows = new Map<string, Arrow>();
  }

  isArrow(node: SceneNode): boolean {
    return node.getPluginData('free-arrow') === 'true' && this.arrows.has(node.id);
  }

  unmarkArrow(node: SceneNode) {
    node.setPluginData('free-arrow', '');
  }

  canDrawArrow(startNode: SceneNode, endNode: SceneNode): boolean {
    const startBounds = NodeUtils.getAbsoluteBounds(startNode);
    const endBounds = NodeUtils.getAbsoluteBounds(endNode);
    const position = GeometryCalculator.getRelativePosition(startBounds, endBounds);

    if (!startBounds || !endBounds) {
      figma.notify('Selected nodes do not have valid bounds.');
      return false;
    }

    if (startNode.id === endNode.id) {
      figma.notify('Cannot create arrow between the same node.');
      return false;
    }

    if (position === 'MM' || position === 'UNKNOWN') {
      figma.notify('The selected objects are too close.');
      return false;
    }

    return true;
  }

  loadArrowsFromPluginData(currentPage: PageNode) {
    try {
      // Clear arrow data from previous page
      this.arrows.clear();
      const savedArrows = currentPage.getPluginData('saved-arrows');
      if (savedArrows && savedArrows.trim() !== '') {
        const arrows = JSON.parse(savedArrows);
        if (Array.isArray(arrows)) {
          arrows.forEach((arrow: Arrow) => {
            if (!this.arrows.has(arrow.id)) {
              this.arrows.set(arrow.id, arrow);
            }
          });
        }
      }
      console.log('[Free Flow] Arrows loaded from: ' + currentPage.name + '. Arrow count:', this.arrows.size);
    } catch (error) {
      console.error('Error loading arrows from plugin data:', error);
      currentPage.setPluginData('saved-arrows', '');
    }
  }

  shouldUpdateArrow(arrowID: string, nodeID: string): boolean {
    if (arrowID === nodeID) {
      return true;
    }

    const arrow = this.arrows.get(arrowID);
    if (!arrow) {
      return false;
    }

    const startNodeId = arrow.startNodeID;
    const endNodeId = arrow.endNodeID;

    if (startNodeId === nodeID || endNodeId === nodeID) {
      return true;
    }

    return false;
  }

  async cleanUpArrows() {

    this.arrows.forEach(async (arrow) => {
      const arrowNode = await figma.getNodeByIdAsync(arrow.id);
      if (!arrowNode) {
        this.removeArrow(arrow.id);

      } else {
        const startNode = await figma.getNodeByIdAsync(arrow.startNodeID);
        const endNode = await figma.getNodeByIdAsync(arrow.endNodeID);
        if (!startNode || !endNode) {
          this.removeArrow(arrow.id);

        }
      }
    });

  }

  async updateArrow(id: string, newStyle?: ArrowStyle, newGeometry?: ArrowGeometry, newText?: ArrowText, newDirection?: ArrowDirection): Promise<VectorNode | null> {

    // If this arrow is already being updated, skip
    if (this.updatingArrows.has(id)) {
      return null;
    }

    const arrow = this.arrows.get(id);
    if (!arrow) {
      console.log('[Free Flow] Arrow not found, probably because it was deleted or updating too fast... we will be fine :)');
      // Clean up the arrow
      this.removeArrow(id);
      return null;
    }

    const arrowNode = await figma.getNodeByIdAsync(id);
    if (!arrowNode) {
      console.log('[Free Flow] Arrow node not found, probably because it was deleted or updating too fast... we will be fine :)');
      // Clean up the arrow
      this.removeArrow(id);
      return null;
    }

    // Mark as updating
    this.updatingArrows.add(id);

    try {

      const [startNode, endNode] = await Promise.all([
        figma.getNodeByIdAsync(arrow.startNodeID),
        figma.getNodeByIdAsync(arrow.endNodeID),
      ]);

      const typedArrowNode = arrowNode as VectorNode;
      const typedStartNode = startNode as SceneNode;
      const typedEndNode = endNode as SceneNode;

      // Get style & geometry (if not provided, use the old style & geometry) 
      const geometry = newGeometry || { ...arrow.geometry };
      const style = newStyle || { ...arrow.style };
      const text = newText || { ...arrow.text } as ArrowText;

      // Get direction, also check if the direction is valid, if not, reassign the direction
      let direction = newDirection || arrow.direction;
      direction = GeometryCalculator.getValidDirection(direction, NodeUtils.getAbsoluteBounds(typedStartNode), NodeUtils.getAbsoluteBounds(typedEndNode));

      await this.updateArrowVector(typedArrowNode, geometry, typedStartNode, typedEndNode, direction);
      this.updateArrowStyle(typedArrowNode, style);
      this.updateArrowName(typedArrowNode, arrow.name);

      // update arrow data
      arrow.geometry = geometry;
      arrow.style = style;
      arrow.name = arrow.name;
      arrow.text = text;
      arrow.position = {
        startRect: NodeUtils.getAbsoluteBounds(typedStartNode),
        endRect: NodeUtils.getAbsoluteBounds(typedEndNode)
      };
      arrow.direction = direction;

      // Refresh arrow text
      await this.refreshArrowText(id, text, typedStartNode, typedEndNode, GeometryCalculator.getAnchorPoints(typedStartNode, typedEndNode, geometry.startMargin, geometry.endMargin, direction));

      return typedArrowNode;
    } catch (error) {
      console.error('Error updating arrow:', error);
      return null;
    } finally {
      // Unmark as updating
      this.updatingArrows.delete(id);
    }
  }


  // Refresh the arrow text node to the latest text status
  async refreshArrowText(arrowID: string, text: ArrowText, startNode: SceneNode, endNode: SceneNode, anchors: ArrowAnchors): Promise<void> {
    const arrow = this.arrows.get(arrowID);
    if (!arrow) {
      console.error('Arrow not found');
      return;
    }

    const textNodeID = arrow.textNodeID;

    if (textNodeID) {
      if (text && text.text && text.text !== '') {
        // If there is text, update the text node
        const textFrameNode = await figma.getNodeByIdAsync(textNodeID) as FrameNode;
        if (textFrameNode) {
          await this.updateArrowText(textFrameNode, text, arrow.style, startNode, endNode, anchors);
          return;
        }
      }
      // If there is no text or text node is not found, remove the text node
      this.removeArrowText(arrowID);
    } else {
      // If there is text, create a new text node
      if (text && text.text && text.text !== '') {
        const textFrameNode = await figma.createFrame();
        await this.updateArrowText(textFrameNode, text, arrow.style, startNode, endNode, anchors);
        arrow.textNodeID = textFrameNode.id;
      }
    }
  }

  async removeArrowText(id: string): Promise<void> {
    try {
      const arrow = this.arrows.get(id);
      if (!arrow) {
        console.error('Arrow not found');
        return;
      }

      const textNodeID = arrow.textNodeID;
      if (textNodeID) {
        const textNode = await figma.getNodeByIdAsync(textNodeID);
        if (textNode) textNode.remove();
      }

      // Update arrow data
      arrow.textNodeID = undefined;
      arrow.text = undefined;
      return;
    } catch (error) {
      console.error('Error removing arrow text:', error);
    }
  }

  isTextInArrow(arrowID: string, textNodeID: string): boolean {
    return this.arrows.get(arrowID)?.textNodeID === textNodeID;
  }

  async removeArrow(id: string): Promise<void> {

    // If this arrow is already being removed, skip
    if (this.removingArrows.has(id)) {
      return;
    }

    this.removingArrows.add(id);

    try {
      const arrowNode = await figma.getNodeByIdAsync(id);
      if (arrowNode) {
        // Remove arrow
        arrowNode.remove();
      }

      // Remove text node
      if (this.arrows.get(id)?.textNodeID) {
        await this.removeArrowText(id);
      }

      this.arrows.delete(id);

    } catch (error) {
      console.error('Error removing arrow:', error);
    } finally {
      // Unmark as removing
      this.removingArrows.delete(id);
    }
  }

  async flipArrow(id: string): Promise<VectorNode | null> {
    const arrow = this.arrows.get(id);
    if (!arrow) {
      console.error('Arrow not found');
      return null;
    }

    // Get a snapshot of the arrow
    const arrowSnapshot = { ...arrow };

    // Update arrow data (flipped)
    arrow.startNodeID = arrowSnapshot.endNodeID;
    arrow.endNodeID = arrowSnapshot.startNodeID;
    arrow.position = { startRect: arrowSnapshot.position.endRect, endRect: arrowSnapshot.position.startRect } as ArrowPosition;
    arrow.name = { startName: arrowSnapshot.name.endName, endName: arrowSnapshot.name.startName } as ArrowName;

    // Update arrow
    const newArrow = await this.updateArrow(id).catch(error => {
      console.error('Error flipping arrow:', error);
      return null;
    });

    return newArrow;
  }

  async updateArrowText(textFrameNode: FrameNode, text: ArrowText, style: ArrowStyle, startNode: SceneNode, endNode: SceneNode, anchors: ArrowAnchors): Promise<void> {
    const textNode = textFrameNode.children[0] as TextNode || figma.createText();
    const { startX, startY, endX, endY } = anchors;
    const backgroundColor = style.color;
    const fontSize = style.weight * 2 + 8;
    const padding = fontSize / 2;
    const cornerRadius = padding;
    await figma.loadFontAsync({ family: 'Inter', style: 'Regular' }).then(() => {
      textNode.fontName = { family: 'Inter', style: 'Regular' };
    });

    // Set text node properties
    textNode.x = 0;
    textNode.y = 0;
    textNode.characters = text.text;
    textNode.fontSize = fontSize;
    textNode.fills = [{ type: 'SOLID', color: ColorUtils.getTextColor(backgroundColor) }];
    textFrameNode.appendChild(textNode);

    // Text frame node properties
    textFrameNode.name = ' ';
    textFrameNode.layoutMode = 'HORIZONTAL';
    textFrameNode.fills = [{ type: 'SOLID', color: ColorUtils.hexToFigmaRGB(backgroundColor) }];
    textFrameNode.layoutSizingVertical = 'HUG';
    textFrameNode.layoutSizingHorizontal = 'HUG';
    textFrameNode.paddingLeft = padding;
    textFrameNode.paddingRight = padding;
    textFrameNode.paddingTop = padding;
    textFrameNode.paddingBottom = padding;
    textFrameNode.cornerRadius = cornerRadius;

    // Center the frame
    textFrameNode.x = (startX + endX - textFrameNode.width) / 2;
    textFrameNode.y = (startY + endY - textFrameNode.height) / 2;

    // Add text to the context
    NodeUtils.appendToContext(textFrameNode, startNode, endNode);
  }

  async createArrow(startNode: SceneNode, endNode: SceneNode, style: ArrowStyle, geometry: ArrowGeometry, text?: ArrowText): Promise<VectorNode> {
    // Validate arrow creation
    if (!this.canDrawArrow(startNode, endNode)) {
      throw new Error('Invalid nodes for arrow creation');
    }

    // Get positions
    const startPos = NodeUtils.getAbsoluteBounds(startNode);
    const endPos = NodeUtils.getAbsoluteBounds(endNode);

    // Get direction
    const direction = GeometryCalculator.getDirection(startPos, endPos);

    // Get names
    const arrowName = { startName: startNode.name, endName: endNode.name } as ArrowName;

    // Create arrow, note: text node is not created, it will be created when the arrow is updated
    const arrow = figma.createVector();
    await this.updateArrowVector(arrow, geometry, startNode, endNode, direction);
    this.updateArrowStyle(arrow, style);
    this.updateArrowName(arrow, arrowName);

    // Mark arrow as Free Flow
    arrow.setPluginData('free-arrow', 'true');

    // Store arrow data
    this.arrows.set(arrow.id, {
      id: arrow.id,
      startNodeID: startNode.id,
      endNodeID: endNode.id,
      position: {
        startRect: startPos,
        endRect: endPos
      },
      geometry: geometry,
      style: style,
      name: arrowName,
      direction: direction
    })

    return arrow;
  }

  async updateArrowVector(arrowNode: VectorNode, geometry: ArrowGeometry, startNode: SceneNode, endNode: SceneNode, direction: ArrowDirection): Promise<void> {
    try {

      arrowNode.x = 0;
      arrowNode.y = 0;

      const startPos = NodeUtils.getAbsoluteBounds(startNode);
      const endPos = NodeUtils.getAbsoluteBounds(endNode);

      const newVertices = GeometryCalculator.getArrowVertices(startPos, endPos, geometry, direction);
      const newSegments = GeometryCalculator.getArrowSegments(newVertices.length, startPos, endPos, geometry, direction);

      await arrowNode.setVectorNetworkAsync({
        vertices: newVertices,
        segments: newSegments
      });

      NodeUtils.appendToContext(arrowNode, startNode, endNode);
    } catch (error) {
      console.error('Error updating arrow geometry:', error);
    }
  }

  updateArrowStyle(arrow: VectorNode, style: ArrowStyle): void {
    const { color, opacity, weight, radius, dash, dashGap } = style;

    arrow.strokeWeight = weight;
    arrow.strokeAlign = 'CENTER';
    arrow.strokes = [{
      type: 'SOLID',
      color: ColorUtils.hexToFigmaRGB(color)
    }];
    arrow.opacity = ColorUtils.numberToFigmaOpacity(opacity);
    arrow.cornerRadius = radius;
    if (dash > 0) {
      arrow.dashPattern = [dash, dashGap];
    } else {
      arrow.dashPattern = [];
    }
  }

  updateArrowName(arrow: VectorNode, name: ArrowName) {
    arrow.name = '[Free Flow] ' + name.startName + ' -> ' + name.endName;
  }

  saveArrowsToPluginData(currentPage: PageNode) {
    const arrows = JSON.stringify(Array.from(this.arrows.values()));
    currentPage.setPluginData('saved-arrows', arrows);
    console.log('[Free Flow] Saved arrows to plugin data:', this.arrows.size);
  }
}

/** UI MESSANGER */
class UIMessenger {
  private isArrowUpToDate: boolean = false; // Flag to track if the arrow is up to date

  constructor(
    private state: State,
    private arrowManager: ArrowManager,
    private selection: Selection
  ) {
    this.initializeUI();
    this.attachMessageListeners();
    this.attachSelectionListeners();
    this.handleClosePlugin();
  }

  initializeUI() {
    figma.showUI(__html__, { width: 240, height: 290, themeColors: true });
    this.sendStateToUI();
  }

  // When the plugin is closed, store the arrow data in plugin data
  handleClosePlugin() {
    figma.on('close', async () => {
      await this.arrowManager.cleanUpArrows();
      this.arrowManager.saveArrowsToPluginData(figma.currentPage);
    });
  }

  sendStateToUI() {
    figma.ui.postMessage({
      type: 'state-update',
      strokeColor: this.state.getState().color,
      strokeOpacity: this.state.getState().opacity,
      strokeStartMargin: this.state.getState().startMargin,
      strokeEndMargin: this.state.getState().endMargin,
      strokeWeight: this.state.getState().weight,
      strokeRadius: this.state.getState().radius,
      strokeDash: this.state.getState().dash,
      strokeDashGap: this.state.getState().dashGap,
      startStrokeCap: this.state.getState().startCap,
      endStrokeCap: this.state.getState().endCap,
      lineStyle: this.state.getState().lineStyle
    })
  }

  async handleStateChange(msg: any) {
    const newState: PluginState = {
      color: msg.strokeColor,
      opacity: msg.strokeOpacity,
      startMargin: msg.strokeStartMargin,
      endMargin: msg.strokeEndMargin,
      weight: msg.strokeWeight,
      radius: msg.strokeRadius,
      dash: msg.strokeDash,
      dashGap: msg.strokeDashGap,
      startCap: msg.startStrokeCap,
      endCap: msg.endStrokeCap,
      lineStyle: msg.lineStyle
    }
    this.state.updateState(newState);

    // If the arrow is selected, update the arrow
    if (this.selection.isArrowSelected()) {
      const arrow = this.selection.getSelectedArrow();
      if (arrow) {
        // Get updated style & geometry
        const newStyle = this.state.getStyle();
        const newGeometry = this.state.getGeometry();
        const newText = { text: msg.text } as ArrowText;

        // Update the arrow
        const newArrow = await this.arrowManager.updateArrow(arrow.id, newStyle, newGeometry, newText);

        // Select the new arrow
        if (newArrow) {
          // Set flag to prevent selection change handler from sending data back
          this.isArrowUpToDate = true;
          figma.currentPage.selection = [newArrow as SceneNode];
        }

        // Reset flag after a delay
        setTimeout(() => {
          this.isArrowUpToDate = false;
        }, 10);
      }
    }
  }

  async handleFlipArrow() {
    const arrow = this.selection.getSelectedArrow();
    if (arrow) {
      const newArrow = await this.arrowManager.flipArrow(arrow.id);
      if (newArrow) {
        figma.currentPage.selection = [newArrow as SceneNode];
        this.sendSelectedArrow(newArrow as SceneNode);
      }
    }
  }

  async handleSetDirection(direction: ArrowDirection) {
    const arrow = this.selection.getSelectedArrow();
    if (arrow) {
      const newArrow = await this.arrowManager.updateArrow(arrow.id, undefined, undefined, undefined, direction).catch(
        error => {
          console.log('Error setting direction:', error);
          figma.notify('Failed to set arrow direction. Please try again.');
          return null;
        }
      )
      if (newArrow) {
        figma.currentPage.selection = [newArrow as SceneNode];
        this.sendSelectedArrow(newArrow as SceneNode);
      }
    }
  }

  async attachMessageListeners() {
    figma.ui.onmessage = async (msg) => {
      try {
        switch (msg.type) {
          case 'draw-arrow':
            this.handleDrawArrow();
            break;
          case 'state-change':
            this.handleStateChange(msg);
            break;
          case 'autodraw-toggle':
            this.selection.setAutoDraw(msg.autoDraw);
            break;
          case 'flip':
            this.handleFlipArrow();
            break;
          case 'set-direction':
            this.handleSetDirection(msg.direction);
            break;
        }
      } catch (error) {
        console.error('Error handling message:', error);
      }
    }
  }

  async attachSelectionListeners() {
    figma.on('selectionchange', () => {

      const currentSelection = figma.currentPage.selection;

      // If the selection is the arrow itself
      if (currentSelection.length === 1 && this.arrowManager.isArrow(currentSelection[0] as SceneNode)) {
        this.selection.setSelectedArrow(currentSelection[0] as VectorNode);
        this.sendSelectedArrow(currentSelection[0]);
        return;
      }

      if (currentSelection.length === 1) {
        this.selection.setFirstSelected(currentSelection[0]);
      }
      else if (currentSelection.length === 2) {
        // If some of the nodes are arrows, return
        if (currentSelection.some(node => this.arrowManager.isArrow(node as SceneNode))) {
          return;
        }

        // Set the first and second selected nodes
        this.selection.setFirstAndSecondSelected(currentSelection);

        // If auto draw is enabled, draw an arrow
        if (this.selection.getAutoDraw()) {
          // Use setTimeout to allow the selection to fully process before drawing
          setTimeout(() => {
            this.handleDrawArrow().catch(error => {
              console.error('Error drawing arrow:', error);
              figma.notify('Failed to auto-draw arrow. Please try again.');
            });
          }, 1);
        }
      }
      else {
        if (this.selection.getSelectedArrow()) {
          figma.ui.postMessage({
            type: 'arrow-deselected'
          })
        }
        this.selection.clearSelectedArrow();
      }

    });
  }

  sendSelectedArrow(arrow: SceneNode) {
    // Skip if the arrow is up to date
    if (this.isArrowUpToDate) {
      return;
    }
    const savedArrow = this.arrowManager.arrows.get(arrow.id);
    const text = savedArrow?.text;
    if (savedArrow) {
      figma.ui.postMessage({
        type: 'arrow-selected',
        startNodeName: savedArrow.name.startName,
        endNodeName: savedArrow.name.endName,
        strokeColor: savedArrow.style.color,
        strokeOpacity: savedArrow.style.opacity,
        strokeStartMargin: savedArrow.geometry.startMargin,
        strokeEndMargin: savedArrow.geometry.endMargin,
        strokeWeight: savedArrow.style.weight,
        startStrokeCap: savedArrow.geometry.startCap,
        endStrokeCap: savedArrow.geometry.endCap,
        lineStyle: savedArrow.geometry.lineStyle,
        text: text?.text || ''
      })
    } else {
      this.arrowManager.unmarkArrow(arrow);
      this.selection.clearSelectedArrow();
    }
  }

  async handleDrawArrow() {
    try {
      const { startNode, endNode } = this.selection.getStartAndEndNodes();
      const style = this.state.getStyle();
      const geometry = this.state.getGeometry();
      const arrow = await this.arrowManager.createArrow(startNode, endNode, style, geometry);

      // Clear selection and select the new arrow
      figma.currentPage.selection = [arrow as SceneNode];
    } catch (error) {
      figma.notify(error instanceof Error ? error.message : 'An error occurred');
    }
  }
}

/** EVENT HANDLER */
class Events {
  private pendingUpdates: Set<string> = new Set();
  private debouncedUpdateArrows: () => void;
  private currentPage: PageNode;

  constructor(
    private arrowManager: ArrowManager,
    private selection: Selection
  ) {
    this.currentPage = figma.currentPage;
    this.arrowManager.loadArrowsFromPluginData(this.currentPage);
    // Create debounced function for updating arrows
    this.debouncedUpdateArrows = DebounceUtils.debounce(() => {
      this.processPendingUpdates();
    }, 100); // 100ms delay

    this.handleAutoUpdate();
    this.handlePageChange();
  }

  private async processPendingUpdates() {
    const arrowIDs = Array.from(this.pendingUpdates);
    this.pendingUpdates.clear();

    // Process updates in parallel for better performance
    // Use a Set to deduplicate arrow IDs and ensure each arrow is only updated once
    const uniqueArrowIDs = new Set(arrowIDs);
    const updatePromises = Array.from(uniqueArrowIDs).map(async (arrowID) => {
      try {
        await this.arrowManager.updateArrow(arrowID);
      } catch (error) {
        console.error(`Error updating arrow ${arrowID}:`, error);
      }
    });

    await Promise.all(updatePromises);
  }

  async handleAutoUpdate() {
    const listener = async (event: NodeChangeEvent) => {
      // Snapshot of current arrows
      const arrowIDs = Array.from(this.arrowManager.arrows.keys());

      for (const change of event.nodeChanges) {
        const { id, type } = change;


        if (type === 'PROPERTY_CHANGE') {
          // Get the changed node and all its descendants
          const changedNode = await figma.getNodeByIdAsync(id);
          let allAffectedIds = [id];
          if (changedNode) {
            allAffectedIds = NodeUtils.getAllDescendantIds(changedNode);
          }
          // Skip changes to the arrows
          if (this.arrowManager.arrows.has(id)) continue;

          // Add affected arrows & descendants to pending updates
          arrowIDs.forEach(arrowID => {
            if (allAffectedIds.some(affectedId => this.arrowManager.shouldUpdateArrow(arrowID, affectedId))) {
              this.pendingUpdates.add(arrowID);
            }
          });
        } else if (type === 'DELETE') {
          // Handle deletions immediately (no debouncing for deletions)
          arrowIDs.forEach(async arrowID => {
            // If the deleted node is arrow text, remove arrow text from arrow
            if (this.arrowManager.isTextInArrow(arrowID, id)) {
              await this.arrowManager.removeArrowText(arrowID);
            }
            // If the deleted node includes an arrow, remove the arrow
            if (this.arrowManager.shouldUpdateArrow(arrowID, id)) {
              await this.arrowManager.removeArrow(arrowID);
            }
          });
        }
      }

      // Trigger debounced update if there are pending updates
      if (this.pendingUpdates.size > 0) {
        this.debouncedUpdateArrows();
      }
    };

    this.currentPage.on('nodechange', listener);
  }

  handlePageChange() {
    figma.on('currentpagechange', () => {
      // clean up arrows from previous page
      this.arrowManager.cleanUpArrows();
      // save arrows to plugin data of previous page
      this.arrowManager.saveArrowsToPluginData(this.currentPage);
      // clear previous page listeners
      this.currentPage.off('nodechange', this.handleAutoUpdate);
      // update current page
      this.currentPage = figma.currentPage;
      // Set up auto-update for the new page
      this.arrowManager.loadArrowsFromPluginData(this.currentPage);
      this.handleAutoUpdate();
    });
  }
}

/** MAIN HANDLER */
class Main {
  private state: State;
  private arrowManager: ArrowManager;
  private selection: Selection;
  private uiMessenger: UIMessenger;
  private event: Events;

  constructor() {
    this.state = new State();
    this.arrowManager = new ArrowManager();
    this.selection = new Selection();
    this.uiMessenger = new UIMessenger(this.state, this.arrowManager, this.selection);
    this.event = new Events(this.arrowManager, this.selection);
  }
}

/** PLUGIN INITIALIZATION */
const plugin = new Main();